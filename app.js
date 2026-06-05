/* =========================================================================
   ĐỐI SOÁT ĐƠN HÀNG — app.js
   Đọc Google Sheet (API key) -> parse future-proof theo TỪ KHÓA ở hàng 2
   -> xử lý merge -> review -> xuất PDF & Excel.
   Chạy 100% trong trình duyệt. Không gửi dữ liệu đi đâu.

   >>> XEM CHANGELOG & PROMPT BÀN GIAO ĐẦY ĐỦ Ở CUỐI FILE index.html <<<
   ========================================================================= */

const APP_VERSION = "1.2.2";

const App = (() => {
  "use strict";

  /* ----- ĐỊNH NGHĨA CỘT (FUTURE-PROOF) -------------------------------------
     Mỗi cột logic = 1 danh sách "matcher". Dò tiêu đề ở HÀNG 2 (index 1).
     Chuẩn hóa: bỏ dấu, thường hóa, gộp khoảng trắng -> so khớp chứa từ khóa.
     Đổi vị trí cột / thêm cột mới -> vẫn nhận đúng nhờ khớp theo nội dung.
     'all' = mọi từ khóa phải xuất hiện; được phép dùng 'not' để loại trừ.   */
  const FIELD_DEFS = [
    { key:"ngayOrder",  label:"Ngày Order",            match:[{all:["ngay","order"]},{all:["ngay","lay","hang"]}] },
    { key:"ngayDV",     label:"Ngày Cung Cấp DV",      match:[{all:["ngay","cung","cap"]},{all:["ngay","dv"]}] },
    { key:"orderId",    label:"Mã Order ID",           match:[{all:["ma","order","id"]}] },
    { key:"maVan",      label:"Mã Đơn Vận/Mã Đơn Hàng",match:[{all:["ma","don","van"]},{all:["ma","don","hang"]}] },
    { key:"sanPham",    label:"Dịch Vụ / Sản Phẩm",    match:[{all:["dich","vu"],not:["vat","tu dong"]},{all:["san","pham"],not:["vat"]}] },
    { key:"soLuong",    label:"Số Lượng",              match:[{all:["so","luong"]}] },
    { key:"donGiaGiam", label:"Đơn Giá (sau giảm)",    match:[{all:["don","gia","co vat","sau giam"],not:["tru sim","truoc"]}] },
    // cột "đơn giá chưa VAT sau giảm & trừ sim trắng" — có thể THIẾU ở bảng cũ
    { key:"donGiaChuaVat", label:"Đơn Giá Chưa VAT (sau giảm, trừ SIM trắng)",
        match:[{all:["don","gia","chua vat","sau giam","tru sim"]},{all:["chua vat","tru sim trang"]}] },
    // các cột phụ trợ để TÍNH khi thiếu cột trên
    { key:"_donGiaTruSim", label:"_đơn giá có VAT sau giảm & trừ sim",
        match:[{all:["don","gia","co vat","sau giam","tru sim"]}], hidden:true },
    { key:"_tongSimTrang", label:"_tổng tiền sim trắng",
        match:[{all:["tong","sim","trang"]},{all:["tien","sim","trang"]}], hidden:true },
    // thông tin xuất hóa đơn (gộp xuống dòng phụ trong PDF)
    { key:"mst",        label:"MST",                   match:[{all:["mst"]},{all:["ma so thue"]}] },
    { key:"tenMua",     label:"Tên Người Mua / Công Ty",match:[{all:["ten","nguoi","mua"]},{all:["ten","cty"]},{all:["ten","cong ty"]}] },
    { key:"diaChi",     label:"Địa Chỉ",               match:[{all:["dia","chi"]}] },
    { key:"email",      label:"Email",                 match:[{all:["email"]}] },
  ];

  // Cột hiển thị trong file xuất, theo thứ tự. ZALO sẽ bỏ 'maVan'.
  // 'stt' = số thứ tự theo ĐƠN (đơn merge nhiều hàng vẫn 1 STT).
  const OUT_COLS = ["stt","ngayOrder","ngayDV","orderId","maVan","sanPham","soLuong","donGiaGiam","donGiaChuaVat"];
  const OUT_LABEL = {
    stt:"STT", ngayOrder:"Ngày Order", ngayDV:"Ngày Cung Cấp DV", orderId:"Mã Order ID",
    maVan:"Mã Đơn Vận / Đơn Hàng", sanPham:"Dịch Vụ / Sản Phẩm", soLuong:"SL",
    donGiaGiam:"Đơn Giá (sau giảm)", donGiaChuaVat:"Đơn Giá Chưa VAT (sau giảm, trừ SIM trắng)"
  };
  const VAT_RATE = 1.1; // 10%

  let state = { months:[], current:0, sheetId:"" };

  /* ----- tiện ích ---------------------------------------------------------*/
  const norm = s => (s==null?"":String(s))
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")   // bỏ dấu tiếng Việt
    .replace(/đ/gi,"d").toLowerCase().replace(/\s+/g," ").trim();
  const matchHeader = (header, defs) => {
    const h = norm(header);
    return defs.some(m =>
      m.all.every(k => h.includes(k)) &&
      (!m.not || m.not.every(k => !h.includes(k)))
    );
  };
  const isBlank = v => v==null || String(v).trim()==="";
  const num = v => {
    if (isBlank(v)) return null;
    const n = Number(String(v).replace(/[.,\s]/g,m=>m===","?"":"")); // bỏ phân tách nghìn
    const cleaned = String(v).replace(/[^\d.\-]/g,"");
    const f = parseFloat(cleaned);
    return isNaN(f)?null:f;
  };
  const fmtNum = n => n==null?"":Math.round(n).toLocaleString("vi-VN");
  const el = (t,c,txt)=>{const e=document.createElement(t);if(c)e.className=c;if(txt!=null)e.textContent=txt;return e;};
  const escapeHtml = s => String(s==null?"":s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));

  /* ----- GIỮ ĐỊNH DẠNG GẠCH NGANG (strikethrough) -------------------------
     Google Sheets cho biết gạch ngang theo 2 cách:
       1) effectiveFormat.textFormat.strikethrough = true  -> gạch cả ô
       2) textFormatRuns[] = các đoạn ký tự, mỗi đoạn có thể strikethrough
     Hàm trả về HTML đã escape, phần gạch bọc trong <s>...</s>.            */
  const richCellHtml = (cell) => {
    if(!cell) return "";
    const text = cell.formattedValue!=null ? String(cell.formattedValue) : "";
    if(text==="") return "";
    const runs = cell.textFormatRuns;
    const wholeStrike = cell.effectiveFormat?.textFormat?.strikethrough===true;
    if(!runs || !runs.length){
      return wholeStrike ? "<s>"+escapeHtml(text)+"</s>" : escapeHtml(text);
    }
    // dựng theo runs: mỗi run bắt đầu tại startIndex (mặc định 0)
    let html="", segs=[];
    for(let i=0;i<runs.length;i++){
      const start=runs[i].startIndex||0;
      const end=(i+1<runs.length)?(runs[i+1].startIndex||text.length):text.length;
      segs.push({start,end,strike:runs[i].format?.strikethrough===true});
    }
    if(segs.length===0 || segs[0].start>0)
      segs.unshift({start:0,end:segs.length?segs[0].start:text.length,strike:wholeStrike});
    segs.forEach(s=>{
      const part=escapeHtml(text.slice(s.start,s.end));
      html += (s.strike? "<s>"+part+"</s>" : part);
    });
    return html;
  };

  // giữ lại splitCodeText (không dùng để tách mã nữa, nhưng còn dùng nơi khác nếu cần)
  const splitCodeText = raw => ({codes:[String(raw||"").trim()].filter(Boolean), text:""});

  return {
    /* =====================================================================
       1) ĐỌC GOOGLE SHEET
       ===================================================================== */
    async load(){
      const url = document.getElementById("sheetUrl").value.trim();
      const key = document.getElementById("apiKey").value.trim();
      const st  = document.getElementById("loadStatus");
      const show=(cls,msg)=>{st.className="status show "+cls;st.innerHTML=msg;};

      const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if(!m){show("err","Link không hợp lệ. Hãy dán đúng link Google Sheet.");return;}
      if(!key){show("err","Chưa nhập API key.");return;}
      const id = m[1]; state.sheetId=id;

      if(document.getElementById("remember").checked)
        try{localStorage.setItem("ds_apikey",key);}catch(e){}

      const btn=document.getElementById("loadBtn");
      btn.disabled=true; show("info",'<span class="spin"></span>Đang đọc danh sách tab...');

      try{
        // lấy metadata: tên tab + merges + màu nền + ĐỊNH DẠNG CHỮ (gạch ngang) + giá trị
        const metaUrl=`https://sheets.googleapis.com/v4/spreadsheets/${id}`+
          `?fields=sheets(properties(title,sheetId),merges,`+
          `data(rowData(values(formattedValue,effectiveFormat(backgroundColor,textFormat(strikethrough)),`+
          `textFormatRuns(startIndex,format(strikethrough))))))`+
          `&includeGridData=true&key=${key}`;
        const res=await fetch(metaUrl);
        if(!res.ok){
          const e=await res.json().catch(()=>({}));
          throw new Error(e.error?.message||("HTTP "+res.status));
        }
        const data=await res.json();
        const tabs=(data.sheets||[]).map(s=>s.properties.title);

        // chỉ giữ tab dạng "T<số> SHOPEE" / "T<số> ZALO"
        const valid=[];
        (data.sheets||[]).forEach(s=>{
          const title=s.properties.title;
          const mt=title.match(/^\s*T\s*(\d{1,2})\s+(SHOPEE|ZALO)\s*$/i);
          if(mt) valid.push({title, month:+mt[1], kind:mt[2].toUpperCase(), raw:s});
        });

        if(!valid.length){
          show("err","Không tìm thấy tab nào dạng <code>T5 SHOPEE</code> / <code>T5 ZALO</code>. "+
            "Các tab đọc được: "+tabs.map(t=>"<code>"+t+"</code>").join(", "));
          btn.disabled=false; return;
        }

        // dựng values + grid TỪ CHÍNH metadata (đã có formattedValue) — không cần gọi values riêng
        for(const v of valid){
          v.grid=v.raw.data?.[0]?.rowData||[];
          v.merges=v.raw.merges||[];
          // values[r][c] = chuỗi hiển thị; lấy từ formattedValue trong grid
          v.values=v.grid.map(rd=>(rd.values||[]).map(c=>c?.formattedValue??""));
        }

        state.months = App._buildMonths(valid);
        state.current = 0;
        App._render();
        show("good",`Đã đọc xong: ${valid.length} tab, gộp thành ${state.months.length} tháng. `+
          `Kiểm tra các mục cảnh báo bên dưới rồi xuất file.`);
      }catch(err){
        show("err","Lỗi: "+err.message+
          "<br><span class='hint'>Kiểm tra: API key đã bật Google Sheets API? Sheet đã để 'Ai có link đều xem'?</span>");
      }finally{ btn.disabled=false; }
    },

    /* =====================================================================
       2) DỰNG DỮ LIỆU THEO THÁNG  (parse + merge + tính toán + phát hiện)
       ===================================================================== */
    _buildMonths(tabs){
      // gom theo tháng
      const byMonth={};
      tabs.forEach(t=>{(byMonth[t.month]=byMonth[t.month]||[]).push(t);});
      const months=[];
      Object.keys(byMonth).map(Number).sort((a,b)=>a-b).forEach(mn=>{
        const sheets=byMonth[mn].map(t=>App._parseSheet(t));
        months.push({month:mn, sheets});
      });
      return months;
    },

    _parseSheet(tab){
      const rows=tab.values;
      const isZalo = tab.kind==="ZALO";
      // ---- map cột: dò hàng 2 (index 1) ----
      const headerRow = rows[1]||[];
      const colMap={}; // key -> col index
      FIELD_DEFS.forEach(def=>{
        for(let c=0;c<headerRow.length;c++){
          if(matchHeader(headerRow[c],def.match)){ colMap[def.key]=c; break; }
        }
      });

      // ---- merge map cho orderId & maVan: cell -> {top, size} ----
      // Google merges: {startRowIndex,endRowIndex,startColumnIndex,endColumnIndex}
      const mergeAnchor={}; // "r_c" -> {rows, anchorRow}
      const mergedInto={};  // "r_c" -> "anchorR_c"
      (tab.merges||[]).forEach(mg=>{
        for(let r=mg.startRowIndex;r<mg.endRowIndex;r++)
          for(let c=mg.startColumnIndex;c<mg.endColumnIndex;c++){
            if(r===mg.startRowIndex&&c===mg.startColumnIndex)
              mergeAnchor[r+"_"+c]={rows:mg.endRowIndex-mg.startRowIndex, c};
            else mergedInto[r+"_"+c]=mg.startRowIndex+"_"+c;
          }
      });

      // ---- màu nền để nhận hàng "ngắt ngày" (xanh) ----
      const isGreenRow = ri=>{
        const rd=tab.grid[ri]; if(!rd||!rd.values)return false;
        // hàng ngắt ngày: nội dung trống + có tô màu (xanh nhạt)
        const firstCells=rd.values.slice(0,8);
        const colored=firstCells.some(c=>{
          const bg=c.effectiveFormat?.backgroundColor; if(!bg)return false;
          const {red=1,green=1,blue=1}=bg;
          // xanh lá nhạt: green trội, không phải trắng
          return green>0.75 && red<0.92 && blue<0.92 && !(red>0.97&&green>0.97&&blue>0.97);
        });
        return colored;
      };

      const out=[];        // hàng dữ liệu chuẩn hóa
      const warnings=[];   // ô mã có chữ ghi chú (để bạn ĐỌC & SỬA full text)
      const emptyDates=[]; // ngày order trống -> review
      const cancelMismatch=[]; // mã có 'hủy' nhưng tên SP chưa có 'hủy'
      const getCell=(ri,key)=>{const c=colMap[key];return c==null?"":(rows[ri]?.[c]??"");};
      // lấy grid cell (để đọc định dạng gạch ngang)
      const getGridCell=(ri,key)=>{const c=colMap[key];if(c==null)return null;
        return tab.grid[ri]?.values?.[c]||null;};
      // html giữ gạch ngang cho 1 cột
      const richHtml=(ri,key)=>richCellHtml(getGridCell(ri,key));

      // "hàng có dữ liệu thật" = có ít nhất 1 trong các cột CỐT LÕI
      // (ngày order, ngày DV, mã order, mã vận, tên SP, số lượng, đơn giá).
      const coreKeys=["ngayOrder","ngayDV","orderId","maVan","sanPham","soLuong","donGiaGiam","donGiaChuaVat"];
      const hasRealData = ri => coreKeys.some(k=>{const c=colMap[k];return c!=null && !isBlank(rows[ri]?.[c]);});

      // tìm hàng dữ liệu thật CUỐI CÙNG -> bỏ mọi hàng trống phía sau
      let lastDataRow = 1; // header ở index 1
      for(let ri=2; ri<rows.length; ri++){ if(hasRealData(ri)) lastDataRow=ri; }

      // data bắt đầu sau header (index 2). Hàng 0 = tổng, hàng 1 = tiêu đề.
      // CHỈ duyệt tới lastDataRow -> không sinh hàng trống thừa ở cuối.
      for(let ri=2; ri<=lastDataRow; ri++){
        const rowEmpty = !hasRealData(ri);

        if(rowEmpty){
          // daybreak chỉ có nghĩa khi nằm GIỮA dữ liệu (đã có data ở trên, và còn data ở dưới
          // — đảm bảo bởi vòng lặp dừng ở lastDataRow). Không thêm 2 daybreak liền nhau.
          if(out.length && out[out.length-1].type==="row")
            out.push({type:"daybreak", ri});
          continue;
        }

        const rec={type:"row", ri, raw:{}, html:{}};
        OUT_COLS.concat(["mst","tenMua","diaChi","email","_donGiaTruSim","_tongSimTrang"]).forEach(k=>{
          rec.raw[k]=getCell(ri,k);
        });
        // HTML giữ gạch ngang cho các cột chữ
        ["orderId","maVan","sanPham"].forEach(k=>{ rec.html[k]=richHtml(ri,k); });

        // ---- YÊU CẦU 4: GIỮ FULL TEXT mã, KHÔNG tách. ----
        // Nếu trong ô mã có chữ (không chỉ là mã thuần) -> đánh dấu để bạn ĐỌC & SỬA cả ô.
        ["orderId","maVan"].forEach(k=>{
          const v=String(rec.raw[k]||"");
          rec[k]={full:v, html:rec.html[k]}; // full text giữ nguyên
          // phát hiện ô có "chữ ghi chú": có chữ cái thường/dấu cách nhiều + không phải chỉ mã
          const hasNote = /[a-zà-ỹ]/i.test(v) && /(hủy|huy|đổi|doi|chờ|cho|lỗi|loi|note|ghi chú|http|còn|con|chưa|chua|\s{2,})/i.test(v)
            || /\n/.test(v);
          if(hasNote && v.trim()!==""){
            warnings.push({ri, field:k, fieldLabel:k==="orderId"?"Mã Order ID":"Mã Đơn Vận/Hàng",
              full:v, keep:v});  // keep = full text để bạn tự sửa
          }
        });

        // merge info cho cột maVan (để xác định block gộp)
        const mvCol=colMap["maVan"], oiCol=colMap["orderId"], odCol=colMap["ngayOrder"];
        rec.mergeTopMaVan = mvCol!=null && mergeAnchor[ri+"_"+mvCol] ? mergeAnchor[ri+"_"+mvCol].rows : 1;
        rec.isMergedSlave = mvCol!=null && !!mergedInto[ri+"_"+mvCol];

        // ngày order có nằm trong block merge không (thừa hưởng ngày phía trên)?
        rec.ngayOrderMerged = odCol!=null && !!mergedInto[ri+"_"+odCol];
        const isContinuation = rec.ngayOrderMerged ||
          (isBlank(rec.raw.orderId) && isBlank(rec.raw.maVan) && out.length &&
           out[out.length-1].type==="row");

        // ngày order trống VÀ không phải dòng nối tiếp -> mới cảnh báo
        if(isBlank(rec.raw.ngayOrder) && !isContinuation){
          emptyDates.push({ri, ngayDV:rec.raw.ngayDV, sanPham:rec.raw.sanPham, fixed:""});
          rec.emptyOrderDate=true;
        }

        // ---- YÊU CẦU 6: ô mã có 'HỦY' nhưng tên SP CHƯA có 'hủy' -> cho sửa tên ----
        const codeText = (String(rec.raw.orderId||"")+" "+String(rec.raw.maVan||""));
        const codeHasHuy = /h[ủu]y/i.test(norm(codeText)) || /huy/i.test(norm(codeText));
        const spHasHuy = /huy/i.test(norm(rec.raw.sanPham||""));
        if(codeHasHuy && !spHasHuy && !isBlank(rec.raw.sanPham)){
          cancelMismatch.push({ri, code:codeText.trim().slice(0,60),
            sanPham:String(rec.raw.sanPham), fixed:String(rec.raw.sanPham)});
          rec.cancelMismatch=true;
        }

        // ---- tính đơn giá chưa VAT nếu thiếu ----
        rec.donGiaGiam = num(rec.raw.donGiaGiam);
        rec.soLuong = num(rec.raw.soLuong);
        let dgChua = num(rec.raw.donGiaChuaVat);
        if(dgChua==null){
          let base = num(rec.raw._donGiaTruSim);
          if(base==null){
            const dgVat=rec.donGiaGiam, simT=num(rec.raw._tongSimTrang), sl=rec.soLuong;
            if(dgVat!=null && dgVat!==0){
              const perSim = (simT!=null && sl) ? simT/Math.abs(sl) : 0;
              base = dgVat - perSim;
            }
          }
          dgChua = (base==null)?null:base/VAT_RATE;
          rec.computedChuaVat=true;
        }
        rec.donGiaChuaVat = dgChua;
        rec.isCancel = (rec.soLuong!=null && rec.soLuong<0);

        // ---- lọc thông tin xuất HĐ: chỉ nhận khi MST hợp lệ HOẶC tên có CHỮ ----
        const mstOk = !isBlank(rec.raw.mst) && /\d{8,}/.test(String(rec.raw.mst).replace(/\D/g,""));
        const tenStr = String(rec.raw.tenMua||"").trim();
        const tenOk = tenStr!=="" && /\p{L}{3,}/u.test(tenStr);
        rec.hasVatInfo = mstOk || tenOk;
        if(!tenOk) rec.raw.tenMua="";
        if(!mstOk) rec.raw.mst="";

        out.push(rec);
      }

      // dọn daybreak thừa ở đầu & cuối (phòng trường hợp biên)
      while(out.length && out[0].type==="daybreak") out.shift();
      while(out.length && out[out.length-1].type==="daybreak") out.pop();

      // ---- SHOPEE: gộp orderId theo block merge của maVan ----
      // Trong block merge maVan, các ô orderId con có thể trống -> dồn FULL TEXT về anchor.
      if(!isZalo){
        for(let i=0;i<out.length;i++){
          const r=out[i]; if(r.type!=="row")continue;
          if(r.mergeTopMaVan>1){
            let span=1;
            const oidParts=r.orderId.full?[r.orderId.full]:[];
            const oidHtml=r.orderId.html?[r.orderId.html]:[];
            for(let j=i+1;j<out.length && span<r.mergeTopMaVan;j++){
              const c=out[j]; if(c.type==="daybreak")continue;
              if(c.isMergedSlave){
                if(c.orderId.full){oidParts.push(c.orderId.full);oidHtml.push(c.orderId.html);}
                c.orderIdHiddenByMerge=true; span++;
              } else break;
            }
            r.orderId.full=oidParts.join("\n");
            r.orderId.html=oidHtml.filter(Boolean).join("<br>");
            r.orderIdRowSpan=r.mergeTopMaVan;
            r.maVanRowSpan=r.mergeTopMaVan;
          }
        }
      }

      // ---- YÊU CẦU 9: đánh STT theo ĐƠN (đơn merge nhiều hàng vẫn 1 STT) ----
      let stt=0;
      for(let i=0;i<out.length;i++){
        const r=out[i]; if(r.type!=="row"){continue;}
        if(r.orderIdHiddenByMerge){ r.stt=null; continue; } // dòng con -> không tăng STT
        stt++; r.stt=stt; r.sttRowSpan=r.orderIdRowSpan||1;
      }

      return {
        title:tab.title, kind:tab.kind, isZalo, month:tab.month,
        colMap, rows:out, warnings, emptyDates, cancelMismatch,
        missingChuaVat: colMap.donGiaChuaVat==null
      };
    },

    /* render + export ở phần kế tiếp (app2.js nối vào) */
    _state:()=>state,
    _util:{norm,num,fmtNum,isBlank,el,escapeHtml,richCellHtml,OUT_COLS,OUT_LABEL,APP_VERSION}
  };
})();

// nạp API key đã lưu
try{const k=localStorage.getItem("ds_apikey");if(k){document.getElementById("apiKey").value=k;
  document.getElementById("remember").checked=true;}}catch(e){}

/* =========================================================================
   PHẦN 2: RENDER MÀN HÌNH REVIEW + PREVIEW, và XUẤT FILE
   ========================================================================= */
(function(){
  const S=()=>App._state();
  const {fmtNum,isBlank,el,OUT_COLS,OUT_LABEL,num}=App._util;

  /* ---------- RENDER ---------- */
  App._render=function(){
    const st=S();
    document.getElementById("emptyCard").style.display="none";
    document.getElementById("workCard").style.display="block";
    document.getElementById("exportBar").style.display="flex";

    // tabbar
    const tb=document.getElementById("tabbar"); tb.innerHTML="";
    const cont=document.getElementById("monthContainer"); cont.innerHTML="";

    st.months.forEach((mo,idx)=>{
      const warnCount=mo.sheets.reduce((a,s)=>a+s.warnings.length+s.emptyDates.length,0);
      const b=el("button",idx===st.current?"active":"","Tháng "+mo.month);
      if(warnCount){const bd=el("span","badge",warnCount);b.appendChild(bd);}
      b.onclick=()=>{st.current=idx;App._render();};
      tb.appendChild(b);

      const block=el("div","month-block"+(idx===st.current?" active":""));
      block.dataset.month=mo.month;
      mo.sheets.forEach(sh=>block.appendChild(App._renderSheet(mo,sh)));
      cont.appendChild(block);
    });

    // summary
    const mo=st.months[st.current];
    const totalRows=mo.sheets.reduce((a,s)=>a+s.rows.filter(r=>r.type==="row").length,0);
    document.getElementById("exportSummary").innerHTML=
      `<span class="pill">Tháng <b>${mo.month}</b></span>`+
      mo.sheets.map(s=>`<span class="pill">${s.kind} <b>${s.rows.filter(r=>r.type==="row").length}</b> dòng</span>`).join("")+
      `<span class="pill">Tổng <b>${totalRows}</b> dòng</span>`;
  };

  App._renderSheet=function(mo,sh){
    const box=el("div");
    box.style.marginBottom="26px";
    const h=el("div");h.style.cssText="font-size:14px;font-weight:700;margin:6px 0 12px;color:#10243f";
    h.textContent=`▸ ${sh.title}  (${sh.kind})`;
    box.appendChild(h);

    if(sh.missingChuaVat){
      const note=el("div","status show info");
      note.innerHTML="ℹ️ Tab này <b>chưa có</b> cột “Đơn giá chưa VAT (sau giảm, trừ SIM trắng)”. "+
        "Hệ thống đã <b>tự tính</b> từ đơn giá có VAT và tiền SIM trắng (VAT 10%).";
      box.appendChild(note);
    }

    // ----- review: text lẫn trong mã -----
    if(sh.warnings.length) box.appendChild(App._panelWarnings(sh));
    // ----- review: ngày order trống -----
    if(sh.emptyDates.length) box.appendChild(App._panelEmpty(sh));
    if(!sh.warnings.length && !sh.emptyDates.length){
      const ok=el("div","review");
      const head=el("div","review-head ok-h");
      head.innerHTML="<span>✓ Không phát hiện text lạ hay ô ngày trống — tab này sạch</span>";
      ok.appendChild(head); box.appendChild(ok);
    }
    // ----- preview bảng -----
    box.appendChild(App._previewTable(sh));
    return box;
  };

  App._renderSheet=function(mo,sh){
    const box=el("div");
    box.style.marginBottom="26px";
    const h=el("div");h.style.cssText="font-size:14px;font-weight:700;margin:6px 0 12px;color:#10243f";
    h.textContent=`▸ ${sh.title}  (${sh.kind})`;
    box.appendChild(h);

    if(sh.missingChuaVat){
      const note=el("div","status show info");
      note.innerHTML="ℹ️ Tab này <b>chưa có</b> cột “Đơn giá chưa VAT (sau giảm, trừ SIM trắng)”. "+
        "Hệ thống đã <b>tự tính</b> từ đơn giá có VAT và tiền SIM trắng (VAT 10%).";
      box.appendChild(note);
    }

    if(sh.warnings.length) box.appendChild(App._panelWarnings(sh));
    if((sh.cancelMismatch||[]).length) box.appendChild(App._panelCancel(sh));
    if(sh.emptyDates.length) box.appendChild(App._panelEmpty(sh));
    if(!sh.warnings.length && !sh.emptyDates.length && !(sh.cancelMismatch||[]).length){
      const ok=el("div","review");
      const head=el("div","review-head ok-h");
      head.innerHTML="<span>✓ Không phát hiện vấn đề cần sửa — tab này sạch</span>";
      ok.appendChild(head); box.appendChild(ok);
    }
    box.appendChild(App._previewTable(sh));
    return box;
  };

  // panel: ô mã có CHỮ GHI CHÚ -> giữ FULL TEXT, bạn tự đọc & sửa (yêu cầu 4)
  App._panelWarnings=function(sh){
    const p=el("div","review");
    const head=el("div","review-head warn-h");
    head.innerHTML=`<span>⚠ ${sh.warnings.length} ô mã có CHỮ GHI CHÚ — đọc & sửa trực tiếp (thuế không nên thấy)</span><span>▾</span>`;
    head.onclick=()=>p.classList.toggle("collapsed");
    const body=el("div","review-body");
    const tbl=el("table","rv");
    tbl.innerHTML="<thead><tr><th>Dòng</th><th>Cột</th><th>Nội dung ô (sửa trực tiếp — giữ nguyên hay xóa chữ tùy bạn)</th></tr></thead>";
    const tb=el("tbody");
    sh.warnings.forEach(w=>{
      const tr=el("tr");
      tr.innerHTML=`<td>${w.ri+1}</td><td>${w.fieldLabel}</td>`;
      const td=el("td");
      const ta=el("textarea");
      ta.value=w.keep; ta.rows=Math.min(4,(String(w.keep).match(/\n/g)||[]).length+1);
      ta.style.cssText="width:100%;font-family:ui-monospace,monospace;font-size:12px;padding:6px 8px;border:1px solid var(--line);border-radius:6px;resize:vertical";
      ta.oninput=e=>{w.keep=e.target.value;App._applyWarningFix(sh,w);};
      td.appendChild(ta);tr.appendChild(td);tb.appendChild(tr);
    });
    tbl.appendChild(tb);body.appendChild(tbl);p.appendChild(head);p.appendChild(body);
    return p;
  };

  App._applyWarningFix=function(sh,w){
    const rec=sh.rows.find(r=>r.ri===w.ri);
    if(!rec)return;
    rec[w.field].full = w.keep;
    rec[w.field].html = App._util.escapeHtml(w.keep).replace(/\n/g,"<br>");
    App._refreshPreview(sh);
  };

  // panel: mã có 'HỦY' nhưng tên SP chưa có 'hủy' (yêu cầu 6)
  App._panelCancel=function(sh){
    const p=el("div","review");
    const head=el("div","review-head warn-h");
    head.innerHTML=`<span>🔁 ${sh.cancelMismatch.length} đơn có chữ HỦY ở ô mã nhưng TÊN SP chưa có chữ “hủy” — sửa tên cho khớp</span><span>▾</span>`;
    head.onclick=()=>p.classList.toggle("collapsed");
    const body=el("div","review-body");
    const tbl=el("table","rv");
    tbl.innerHTML="<thead><tr><th>Dòng</th><th>Ô mã (có chữ hủy)</th><th>Sửa lại TÊN sản phẩm/dịch vụ</th></tr></thead>";
    const tb=el("tbody");
    sh.cancelMismatch.forEach(cm=>{
      const tr=el("tr");
      tr.innerHTML=`<td>${cm.ri+1}</td><td><span class="tag txt">${App._util.escapeHtml(cm.code)}</span></td>`;
      const td=el("td");const inp=el("input");inp.type="text";inp.value=cm.fixed;
      inp.style.fontSize="12px";
      inp.oninput=e=>{cm.fixed=e.target.value;const rec=sh.rows.find(r=>r.ri===cm.ri);
        if(rec){rec.raw.sanPham=e.target.value;rec.html.sanPham=App._util.escapeHtml(e.target.value);}
        App._refreshPreview(sh);};
      td.appendChild(inp);tr.appendChild(td);tb.appendChild(tr);
    });
    tbl.appendChild(tb);body.appendChild(tbl);p.appendChild(head);p.appendChild(body);
    return p;
  };

  App._panelEmpty=function(sh){
    const p=el("div","review");
    const head=el("div","review-head empty-h");
    head.innerHTML=`<span>● ${sh.emptyDates.length} dòng TRỐNG ngày order — điền để sửa luôn</span><span>▾</span>`;
    head.onclick=()=>p.classList.toggle("collapsed");
    const body=el("div","review-body");
    const tbl=el("table","rv");
    tbl.innerHTML="<thead><tr><th>Dòng</th><th>Ngày cung cấp DV</th><th>Sản phẩm</th><th>Điền ngày order</th></tr></thead>";
    const tb=el("tbody");
    sh.emptyDates.forEach(ed=>{
      const tr=el("tr");
      tr.innerHTML=`<td>${ed.ri+1}</td><td>${ed.ngayDV||'<span class=muted>—</span>'}</td>`+
        `<td style="max-width:340px;white-space:normal">${App._util.escapeHtml(ed.sanPham||'')}</td>`;
      const td=el("td");const inp=el("input");inp.type="text";inp.placeholder="vd 5/4/2025";inp.value=ed.fixed;
      inp.oninput=e=>{ed.fixed=e.target.value;const rec=sh.rows.find(r=>r.ri===ed.ri);
        if(rec)rec.raw.ngayOrder=e.target.value;App._refreshPreview(sh);};
      td.appendChild(inp);tr.appendChild(td);tb.appendChild(tr);
    });
    tbl.appendChild(tb);body.appendChild(tbl);p.appendChild(head);p.appendChild(body);
    return p;
  };

  App._previewTable=function(sh){
    const wrap=el("div","preview-scroll");
    wrap.dataset.sheet=sh.title;
    wrap.appendChild(App._buildPreviewInner(sh));
    const lg=el("div","legend");
    lg.innerHTML=
      '<span><span class="sw" style="background:var(--day-break)"></span>Ngắt sang ngày mới</span>'+
      '<span><span class="sw" style="background:var(--cancel)"></span>Đơn hủy (SL âm)</span>'+
      '<span><span class="sw" style="background:var(--vat-row)"></span>Dòng thông tin xuất hóa đơn</span>';
    const cell=el("div");cell.appendChild(wrap);cell.appendChild(lg);
    return cell;
  };

  App._refreshPreview=function(sh){
    document.querySelectorAll('.preview-scroll').forEach(w=>{
      if(w.dataset.sheet===sh.title){w.innerHTML="";w.appendChild(App._buildPreviewInner(sh));}
    });
  };

  App._buildPreviewInner=function(sh){
    const cols=OUT_COLS.filter(c=>!(sh.isZalo&&c==="maVan"));
    const t=el("table","pv");
    let thead="<thead><tr>";
    cols.forEach(c=>thead+=`<th>${OUT_LABEL[c]}</th>`);
    thead+="</tr></thead>";
    t.innerHTML=thead;
    const tb=el("tbody");

    sh.rows.forEach(r=>{
      if(r.type==="daybreak"){
        const tr=el("tr","daybreak");tr.innerHTML=`<td colspan="${cols.length}"></td>`;tb.appendChild(tr);return;
      }
      const tr=el("tr",r.isCancel?"cancel":"");
      cols.forEach(c=>{
        // STT + cột mã có rowspan ở SHOPEE (gộp theo đơn)
        if((c==="stt"||c==="orderId"||c==="maVan") && !sh.isZalo){
          if(r.orderIdHiddenByMerge) return; // bị anchor span phủ
          const td=el("td", c==="stt"?"num":"code merged-cell");
          if(r.orderIdRowSpan) td.rowSpan=r.orderIdRowSpan;
          if(c==="stt") td.textContent = r.stt==null?"":r.stt;
          else if(c==="orderId") td.innerHTML=r.orderId.html||"";
          else td.innerHTML=r.maVan.html||"";
          tr.appendChild(td);return;
        }
        const td=el("td");
        if(c==="stt"){td.className="num";td.textContent=r.stt==null?"":r.stt;tr.appendChild(td);return;}
        if(c==="orderId"){td.className="code";td.innerHTML=r.orderId.html||"";tr.appendChild(td);return;}
        if(c==="maVan"){td.className="code";td.innerHTML=r.maVan.html||"";tr.appendChild(td);return;}
        if(c==="sanPham"){td.innerHTML=r.html.sanPham||App._util.escapeHtml(r.raw.sanPham||"");tr.appendChild(td);return;}
        let v="";
        if(c==="donGiaGiam"){v=fmtNum(r.donGiaGiam);td.className="num";}
        else if(c==="donGiaChuaVat"){v=fmtNum(r.donGiaChuaVat);td.className="num";}
        else if(c==="soLuong"){v=r.soLuong==null?"":r.soLuong;td.className="num";}
        else v=r.raw[c]??"";
        td.textContent=v;
        tr.appendChild(td);
      });
      tb.appendChild(tr);

      if(r.hasVatInfo){
        const tr2=el("tr","vatinfo");
        const parts=[];
        if(!isBlank(r.raw.mst))parts.push("MST: "+r.raw.mst);
        if(!isBlank(r.raw.tenMua))parts.push("Tên: "+r.raw.tenMua);
        if(!isBlank(r.raw.diaChi))parts.push("Đ/c: "+r.raw.diaChi);
        if(!isBlank(r.raw.email))parts.push("Email: "+r.raw.email);
        tr2.innerHTML=`<td colspan="${cols.length}">↳ ${App._util.escapeHtml(parts.join("  |  "))}</td>`;
        tb.appendChild(tr2);
      }
    });
    t.appendChild(tb);
    return t;
  };

  App._util.OUT_COLS_for=sh=>OUT_COLS.filter(c=>!(sh.isZalo&&c==="maVan"));
})();

/* =========================================================================
   PHẦN 3: XUẤT EXCEL & PDF
   ========================================================================= */
(function(){
  const S=()=>App._state();
  const {fmtNum,isBlank,OUT_LABEL,num}=App._util;
  const COLW={stt:4,ngayOrder:11,ngayDV:11,orderId:20,maVan:20,sanPham:44,soLuong:5,donGiaGiam:14,donGiaChuaVat:18};

  function monthsToExport(scope){
    const st=S();
    return scope==="all"?st.months:[st.months[st.current]];
  }
  function titleFor(mo){
    let yr="";
    for(const sh of mo.sheets){for(const r of sh.rows){if(r.type==="row"){
      const m=String(r.raw.ngayOrder||r.raw.ngayDV||"").match(/(20\d{2})/);if(m){yr=m[1];break;}}}if(yr)break;}
    return `Bảng đối soát đơn hàng, theo dõi bán hàng - Tháng ${mo.month}/${yr||"20XX"}`;
  }
  function yearOf(mo){return (titleFor(mo).match(/20\d{2}/)||["20XX"])[0];}

  /* ---------- chuẩn bị ma trận dữ liệu 1 sheet để xuất ----------
     Trả về aoa (giá trị text cho Excel) + htmlMx (HTML giữ gạch ngang cho PDF)
     + merges + rowMeta. */
  function sheetMatrix(sh){
    const cols=App._util.OUT_COLS_for(sh);
    const header=cols.map(c=>OUT_LABEL[c]);
    const aoa=[header];
    const htmlMx=[header.map(h=>App._util.escapeHtml(h))];
    const merges=[];
    const rowMeta=["header"];

    sh.rows.forEach(r=>{
      if(r.type==="daybreak"){aoa.push(cols.map(()=> ""));htmlMx.push(cols.map(()=> ""));rowMeta.push("daybreak");return;}
      const slave = r.orderIdHiddenByMerge && !sh.isZalo;
      const line=cols.map(c=>cellVal(r,c,sh,slave));
      const lineH=cols.map(c=>cellHtml(r,c,sh,slave));
      aoa.push(line);htmlMx.push(lineH);rowMeta.push(r.isCancel?"cancel":"data");

      // merge cho block anchor (SHOPEE): STT + orderId + maVan
      if(!slave && r.orderIdRowSpan && r.orderIdRowSpan>1 && !sh.isZalo){
        const top=aoa.length-1;
        ["stt","orderId","maVan"].forEach(k=>{
          const ci=cols.indexOf(k);
          if(ci>=0) merges.push({s:{r:top,c:ci},e:{r:top+r.orderIdRowSpan-1,c:ci}});
        });
      }
      if(r.hasVatInfo){
        const parts=[];
        if(!isBlank(r.raw.mst))parts.push("MST: "+r.raw.mst);
        if(!isBlank(r.raw.tenMua))parts.push("Tên: "+r.raw.tenMua);
        if(!isBlank(r.raw.diaChi))parts.push("Đ/c: "+r.raw.diaChi);
        if(!isBlank(r.raw.email))parts.push("Email: "+r.raw.email);
        const txt="↳ "+parts.join("   |   ");
        const vrow=cols.map(()=> "");vrow[0]=txt;
        const vrowH=cols.map(()=> "");vrowH[0]=App._util.escapeHtml(txt);
        aoa.push(vrow);htmlMx.push(vrowH);rowMeta.push("vatinfo");
        merges.push({s:{r:aoa.length-1,c:0},e:{r:aoa.length-1,c:cols.length-1}});
      }
    });
    return {cols,aoa,htmlMx,merges,rowMeta};
  }
  // giá trị text (Excel)
  function cellVal(r,c,sh,slave){
    if(c==="stt"){ if(slave)return ""; return r.stt==null?"":r.stt; }
    if(c==="orderId"){ if(slave)return ""; return r.orderId.full||""; }
    if(c==="maVan"){ if(slave)return ""; return r.maVan.full||""; }
    if(c==="sanPham")return r.raw.sanPham??"";
    if(c==="donGiaGiam")return r.donGiaGiam==null?"":r.donGiaGiam;
    if(c==="donGiaChuaVat")return r.donGiaChuaVat==null?"":Math.round(r.donGiaChuaVat);
    if(c==="soLuong")return r.soLuong==null?"":r.soLuong;
    return r.raw[c]??"";
  }
  // HTML giữ gạch ngang (PDF)
  function cellHtml(r,c,sh,slave){
    if(c==="stt"){ if(slave)return ""; return r.stt==null?"":String(r.stt); }
    if(c==="orderId"){ if(slave)return ""; return r.orderId.html||""; }
    if(c==="maVan"){ if(slave)return ""; return r.maVan.html||""; }
    if(c==="sanPham")return r.html?.sanPham||App._util.escapeHtml(r.raw.sanPham||"");
    if(c==="donGiaGiam")return r.donGiaGiam==null?"":Number(r.donGiaGiam).toLocaleString("vi-VN");
    if(c==="donGiaChuaVat")return r.donGiaChuaVat==null?"":Math.round(r.donGiaChuaVat).toLocaleString("vi-VN");
    if(c==="soLuong")return r.soLuong==null?"":String(r.soLuong);
    return App._util.escapeHtml(r.raw[c]??"");
  }

  /* ---------- EXCEL (có style: màu nền, viền, số thật) ---------- */
  const BORDER={top:{style:"thin",color:{rgb:"BBBBBB"}},bottom:{style:"thin",color:{rgb:"BBBBBB"}},
                left:{style:"thin",color:{rgb:"BBBBBB"}},right:{style:"thin",color:{rgb:"BBBBBB"}}};
  const numFmt='#,##0';
  App.exportExcel=function(scope){
    const mos=monthsToExport(scope);
    mos.forEach(mo=>{
      const wb=XLSX.utils.book_new();
      mo.sheets.forEach(sh=>{
        const {cols,aoa,merges,rowMeta}=sheetMatrix(sh);
        const title=titleFor(mo);
        // chèn dòng tiêu đề lớn (đẩy mọi thứ xuống 1 dòng)
        const data=[cols.map((_,i)=> i===0?title:"")].concat(aoa);
        const meta=["title"].concat(rowMeta);

        const ws={};
        const moneyCols=new Set([cols.indexOf("soLuong"),cols.indexOf("donGiaGiam"),cols.indexOf("donGiaChuaVat")]);
        const sttCol=cols.indexOf("stt");
        const range={s:{r:0,c:0},e:{r:data.length-1,c:cols.length-1}};
        for(let r=0;r<data.length;r++){
          for(let c=0;c<cols.length;c++){
            const addr=XLSX.utils.encode_cell({r,c});
            let v=data[r][c];
            const m=meta[r];
            const cell={};
            const isMoney=moneyCols.has(c), isStt=(c===sttCol);
            if((isMoney||isStt) && m!=="title" && m!=="header" && m!=="vatinfo" && v!=="" && v!=null && !isNaN(Number(v))){
              cell.t="n"; cell.v=Number(v); if(isMoney)cell.z=numFmt;
            } else {
              cell.t="s"; cell.v=(v==null?"":String(v));
            }
            const st={border:BORDER,alignment:{vertical:"top",wrapText:true}};
            if(m==="title"){st.font={bold:true,sz:13};st.alignment={horizontal:"center",vertical:"center"};delete st.border;}
            else if(m==="header"){st.font={bold:true,color:{rgb:"FFFFFF"},sz:9};
              st.fill={fgColor:{rgb:"10243F"}};st.alignment={horizontal:"center",vertical:"center",wrapText:true};}
            else if(m==="daybreak"){st.fill={fgColor:{rgb:"CFE8D8"}};}
            else if(m==="vatinfo"){st.fill={fgColor:{rgb:"FFF7E6"}};st.font={italic:true,sz:8,color:{rgb:"5A4200"}};}
            else if(m==="cancel"){st.fill={fgColor:{rgb:"FDECEA"}};}
            if(isMoney&&m!=="title"&&m!=="header") st.alignment={...st.alignment,horizontal:"right"};
            if(isStt&&m!=="title"&&m!=="header") st.alignment={...st.alignment,horizontal:"center"};
            cell.s=st;
            ws[addr]=cell;
          }
        }
        ws["!ref"]=XLSX.utils.encode_range(range);
        ws["!cols"]=cols.map(c=>({wch:COLW[c]||12}));
        // chiều cao dòng ngắt ngày mỏng
        ws["!rows"]=meta.map(m=> m==="daybreak"?{hpt:6}:(m==="title"?{hpt:22}:{}));
        // merges: title + (merges dịch xuống 1 do chèn title)
        const allMerges=[{s:{r:0,c:0},e:{r:0,c:cols.length-1}}];
        merges.forEach(m=>allMerges.push({s:{r:m.s.r+1,c:m.s.c},e:{r:m.e.r+1,c:m.e.c}}));
        ws["!merges"]=allMerges;
        XLSX.utils.book_append_sheet(wb,ws,("T"+mo.month+" "+sh.kind).slice(0,31));
      });
      const yr=(titleFor(mo).match(/20\d{2}/)||["20XX"])[0];
      XLSX.writeFile(wb,`Doi-soat_T${mo.month}_${yr}.xlsx`);
    });
  };

  /* ---------- PDF: IN RIÊNG từng bảng (SHOPEE / ZALO) ----------
     scope: 'current' | 'all'.  kind: 'SHOPEE' | 'ZALO' | null(cả hai, nhưng mỗi
     bảng một cửa sổ in riêng).  Mỗi bảng = 1 cửa sổ in -> 1 file PDF riêng.   */
  App.exportPDF=function(scope,kind){
    const mos=monthsToExport(scope);
    // gom danh sách (tháng, sheet) cần in
    const jobs=[];
    mos.forEach(mo=>mo.sheets.forEach(sh=>{
      if(!kind || sh.kind===kind) jobs.push({mo,sh});
    }));
    if(!jobs.length)return;
    // mở từng cửa sổ in tuần tự (popup nhiều cùng lúc dễ bị chặn -> mở lần lượt)
    let idx=0;
    const openNext=()=>{
      if(idx>=jobs.length)return;
      const {mo,sh}=jobs[idx++];
      const html=buildPdfHtml(mo,sh);
      const w=window.open("","_print_"+Date.now());
      if(!w){alert("Trình duyệt chặn cửa sổ in. Hãy cho phép pop-up cho trang này rồi thử lại.");return;}
      w.document.write(html);w.document.close();
      w.onload=()=>{ setTimeout(()=>{ w.print();
        // sau khi in xong (hoặc đóng), mở bảng kế tiếp
        setTimeout(openNext, 800);
      },300); };
    };
    openNext();
  };

  function buildPdfHtml(mo,sh){
    const {cols,htmlMx,rowMeta,merges}=sheetMatrix(sh);
    let html=`<html><head><meta charset="utf-8"><title>${titleFor(mo)} — ${sh.kind}</title><style>
      @page{size:A4 landscape;margin:8mm}
      *{box-sizing:border-box}
      html,body{margin:0;padding:0}
      body{font-family:"Times New Roman",serif;color:#000}
      .doc-title{text-align:center;font-size:14pt;font-weight:bold;margin:0 0 4px}
      .sheet-title{text-align:center;font-size:10.5pt;font-weight:bold;margin:0 0 8px;color:#222}
      table{border-collapse:collapse;width:100%;table-layout:fixed}
      th,td{border:0.6pt solid #333;padding:2px 3px;font-size:7pt;word-wrap:break-word;
        overflow-wrap:break-word;vertical-align:top;line-height:1.2}
      th{background:#0c1c33 !important;color:#ffffff !important;font-size:6.9pt;font-weight:bold;
        text-align:center;-webkit-print-color-adjust:exact;print-color-adjust:exact;
        border-color:#0c1c33}
      td.num{text-align:right;font-variant-numeric:tabular-nums}
      td.stt{text-align:center}
      td.code{font-family:"Consolas","Courier New",monospace;font-size:6.4pt;word-break:break-all}
      s{color:#b00020}
      tr.daybreak td{background:#bfe3cc !important;height:3px;padding:0;
        -webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr.cancel td{background:#fde0dd !important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr.vatinfo td{background:#fff3da !important;font-size:6.2pt;font-style:italic;color:#5a4200;
        text-align:left;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr{page-break-inside:avoid}
      thead{display:table-header-group}
    </style></head><body>`;
    html+=`<div class="doc-title">${esc(titleFor(mo))}</div>`;
    html+=`<div class="sheet-title">${esc(sh.title)} — ${sh.kind}</div>`;

    const spanAt={}, skip={};
    merges.forEach(m=>{
      spanAt[m.s.r+"_"+m.s.c]=m.e.r-m.s.r+1;
      for(let rr=m.s.r+1;rr<=m.e.r;rr++) skip[rr+"_"+m.s.c]=1;
    });
    const totalW=cols.reduce((a,c)=>a+(COLW[c]||12),0);
    html+="<table><colgroup>"+cols.map(c=>`<col style="width:${((COLW[c]||12)/totalW*100).toFixed(2)}%">`).join("")+"</colgroup>";
    html+="<thead><tr>"+cols.map(c=>`<th>${esc(OUT_LABEL[c])}</th>`).join("")+"</tr></thead><tbody>";
    for(let ri=0;ri<htmlMx.length;ri++){
      const meta=rowMeta[ri];
      if(meta==="header")continue;
      if(meta==="daybreak"){html+=`<tr class="daybreak"><td colspan="${cols.length}"></td></tr>`;continue;}
      if(meta==="vatinfo"){html+=`<tr class="vatinfo"><td colspan="${cols.length}">${htmlMx[ri][0]}</td></tr>`;continue;}
      html+=`<tr class="${meta==='cancel'?'cancel':''}">`;
      for(let ci=0;ci<cols.length;ci++){
        if(skip[ri+"_"+ci])continue;
        const c=cols[ci];
        const isNum=(c==="donGiaGiam"||c==="donGiaChuaVat"||c==="soLuong");
        const isStt=(c==="stt"); const isCode=(c==="orderId"||c==="maVan");
        const sp=spanAt[ri+"_"+ci]?` rowspan="${spanAt[ri+"_"+ci]}"`:"";
        const cls=(isNum?"num":"")+(isStt?" stt":"")+(isCode?" code":"");
        html+=`<td class="${cls.trim()}"${sp}>${htmlMx[ri][ci]}</td>`;
      }
      html+="</tr>";
    }
    html+="</tbody></table></body></html>";
    return html;
  }

  function esc(s){return String(s==null?"":s).replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]));}
})();

/* =========================================================================
   PHẦN 4: MỞ LẠI FILE EXCEL ĐÃ XUẤT để xem/sửa
   ========================================================================= */
(function(){
  const dz=document.getElementById("dropZone");
  const inp=document.getElementById("xlsxInput");
  if(!dz)return;
  dz.onclick=()=>inp.click();
  dz.ondragover=e=>{e.preventDefault();dz.style.borderColor="var(--brand)";dz.style.background="#f0f6ff";};
  dz.ondragleave=()=>{dz.style.borderColor="";dz.style.background="";};
  dz.ondrop=e=>{e.preventDefault();dz.style.borderColor="";dz.style.background="";
    if(e.dataTransfer.files[0])readXlsx(e.dataTransfer.files[0]);};
  inp.onchange=e=>{if(e.target.files[0])readXlsx(e.target.files[0]);};

  function readXlsx(file){
    const st=document.getElementById("loadStatus");
    st.className="status show info";st.innerHTML='<span class="spin"></span>Đang đọc '+file.name+'...';
    const fr=new FileReader();
    fr.onload=ev=>{
      try{
        const wb=XLSX.read(new Uint8Array(ev.target.result),{type:"array"});
        const months=App._monthsFromWorkbook(wb);
        if(!months.length){st.className="status show err";
          st.innerHTML="Không nhận ra tab dạng <code>T7 ZALO</code>/<code>T7 SHOPEE</code> trong file này.";return;}
        App._state().months=months; App._state().current=0; App._render();
        st.className="status show good";
        st.innerHTML=`Đã mở lại file: ${months.length} tháng. Bạn có thể sửa rồi xuất lại.`;
      }catch(err){st.className="status show err";st.innerHTML="Không đọc được file: "+err.message;}
    };
    fr.readAsArrayBuffer(file);
  }

  // dựng state.months từ workbook đã xuất (cấu trúc cố định: r0 title, r1 header, r2+ data)
  App._monthsFromWorkbook=function(wb){
    const {num}=App._util;
    const byMonth={};
    wb.SheetNames.forEach(name=>{
      const mt=name.match(/^\s*T\s*(\d{1,2})\s+(SHOPEE|ZALO)\s*$/i);
      if(!mt)return;
      const month=+mt[1], kind=mt[2].toUpperCase(), isZalo=kind==="ZALO";
      const ws=wb.Sheets[name];
      const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:""});
      const header=aoa[1]||[];
      // map cột theo nhãn xuất ra
      const find=(...kw)=>header.findIndex(h=>{const s=String(h).toLowerCase();
        return kw.every(k=>s.includes(k));});
      const ci={
        stt:find("stt"),
        ngayOrder:find("ngày","order")>=0?find("ngày","order"):find("order"),
        ngayDV:find("cung","cấp"), orderId:find("order","id"),
        maVan:isZalo?-1:find("đơn","vận"),
        sanPham:find("dịch","vụ")>=0?find("dịch","vụ"):find("sản","phẩm"),
        soLuong:find("sl")>=0?find("sl"):find("số","lượng"),
        donGiaGiam:header.findIndex(h=>/sau giảm/i.test(h)&&!/chưa vat/i.test(h)),
        donGiaChuaVat:find("chưa","vat")
      };
      const mergeRows={};
      (ws["!merges"]||[]).forEach(m=>{ if(m.e.r>m.s.r) mergeRows[m.s.r]=Math.max(mergeRows[m.s.r]||1,m.e.r-m.s.r+1); });

      const esc=App._util.escapeHtml;
      const rows=[]; let stt=0;
      for(let r=2;r<aoa.length;r++){
        const row=aoa[r]; if(!row)continue;
        const allEmpty=row.every(v=>String(v).trim()==="");
        if(String(row[0]).trim().startsWith("↳")){
          const prev=rows[rows.length-1];
          if(prev){prev.hasVatInfo=true;
            // tách lại MST/Tên/Đc/Email từ dòng phụ
            const line=String(row[0]).replace(/^↳\s*/,"");
            line.split(/\s*\|\s*/).forEach(p=>{
              const mm=p.match(/^(MST|Tên|Đ\/c|Email):\s*(.*)$/i);
              if(mm){const k={mst:"mst","tên":"tenMua","đ/c":"diaChi",email:"email"}[mm[1].toLowerCase()];
                if(k)prev.raw[k]=mm[2];}
            });
          }
          continue;
        }
        if(allEmpty){
          if(rows.length && rows[rows.length-1].type==="row") rows.push({type:"daybreak",ri:r});
          continue;
        }
        const g=k=>{const i=ci[k];return (i==null||i<0)?"":(row[i]??"");};
        const oid=String(g("orderId")), mv=String(g("maVan")), sp=String(g("sanPham"));
        const rec={type:"row",ri:r,raw:{
          ngayOrder:g("ngayOrder"),ngayDV:g("ngayDV"),sanPham:sp,mst:"",tenMua:"",diaChi:"",email:""},
          html:{sanPham:esc(sp)},
          orderId:{full:oid,html:esc(oid).replace(/\n/g,"<br>")},
          maVan:{full:mv,html:esc(mv).replace(/\n/g,"<br>")},
          soLuong:num(g("soLuong")),donGiaGiam:num(g("donGiaGiam")),
          donGiaChuaVat:num(g("donGiaChuaVat"))
        };
        rec.isCancel=(rec.soLuong!=null&&rec.soLuong<0);
        if(mergeRows[r]){rec.orderIdRowSpan=mergeRows[r];rec.maVanRowSpan=mergeRows[r];}
        stt++; rec.stt=ci.stt>=0&&g("stt")!==""?num(g("stt")):stt; rec.sttRowSpan=mergeRows[r]||1;
        rows.push(rec);
      }
      // dọn daybreak thừa đầu/cuối
      while(rows.length && rows[0].type==="daybreak") rows.shift();
      while(rows.length && rows[rows.length-1].type==="daybreak") rows.pop();
      const sh={title:name,kind,isZalo,month,rows,warnings:[],emptyDates:[],cancelMismatch:[],
        missingChuaVat:false,colMap:{}};
      (byMonth[month]=byMonth[month]||[]).push(sh);
    });
    const months=[];
    Object.keys(byMonth).map(Number).sort((a,b)=>a-b).forEach(m=>months.push({month:m,sheets:byMonth[m]}));
    return months;
  };
})();

/* =========================================================================
   CHANGELOG (hiển thị trong web qua mục "Changelog & Ghi chú kỹ thuật")
   Mỗi lần sửa: tăng APP_VERSION (đầu file) và thêm 1 mục ở ĐẦU danh sách.
   ========================================================================= */
const CHANGELOG_HTML = `
<b>v1.2.2</b> — (bản hiện tại)
<ul style="margin:4px 0 10px">
  <li><b>Bỏ hết hàng trống thừa</b> ở cuối bảng: chỉ duyệt tới hàng dữ liệu thật
      cuối cùng; dải "ngắt ngày" chỉ giữ khi nằm GIỮA dữ liệu.</li>
</ul>
<b>v1.2.1</b>
<ul style="margin:4px 0 10px">
  <li>Thêm <b>favicon SIMGLOBE</b> trên thanh địa chỉ / tab trình duyệt.</li>
  <li>Đổi tiêu đề tab thành "SIMGLOBE v… · Đối Soát Đơn Hàng".</li>
  <li>Thêm <b>cache-busting</b> (<code>app.js?v=…</code>) để trình duyệt luôn tải bản mới,
      tránh kẹt version cũ.</li>
</ul>
<b>v1.2.0</b>
<ul style="margin:4px 0 10px">
  <li>Thêm <b>số phiên bản</b> hiển thị trên thanh tiêu đề + changelog này.</li>
  <li>Thêm <b>logo SIMGLOBE</b> trên header (nhúng sẵn, không cần file ngoài).</li>
  <li>Thêm cột <b>STT</b> đầu bảng, đánh số theo ĐƠN (đơn gộp nhiều hàng = 1 STT).</li>
  <li>Ô mã <code>b...</code> giờ <b>giữ nguyên full text</b> để tự đọc &amp; sửa
      (không tự tách mã/ghi chú nữa).</li>
  <li>Giữ <b>định dạng gạch ngang</b> (strikethrough) từ Google Sheet khi xem &amp; xuất.</li>
  <li>Phát hiện đơn có chữ <b>HỦY</b> ở ô mã nhưng tên SP chưa có "hủy" → cho sửa tên.</li>
  <li><b>PDF in RIÊNG</b> bảng SHOPEE và ZALO (không in chung).</li>
  <li>Header PDF <b>đậm màu hơn</b> để in rõ.</li>
</ul>
<b>v1.1.0</b>
<ul style="margin:4px 0 10px">
  <li>Sửa cảnh báo "ngày order trống" báo nhầm dòng con của đơn gộp.</li>
  <li>Lọc dòng phụ hóa đơn rác (tên chỉ là số).</li>
  <li>Sửa PDF bị trang trắng; ngắt trang sạch.</li>
  <li>Excel: số là số thật + màu nền (ngắt ngày, đơn hủy, dòng HĐ).</li>
  <li>Thêm: thả ngược file Excel đã xuất vào để xem/sửa.</li>
</ul>
<b>v1.0.0</b>
<ul style="margin:4px 0 0">
  <li>Bản đầu: đọc Google Sheet qua API key, nhận tab T# SHOPEE/ZALO, gộp theo
      tháng, review (mã có ghi chú, ngày trống), xuất PDF (ngang) + Excel,
      xử lý merge SHOPEE, ngắt ngày, đơn hủy, dòng phụ hóa đơn.</li>
</ul>
<div style="margin-top:10px;padding-top:8px;border-top:1px solid #e2e7f0;color:#667">
  Ghi chú: cột nhận theo <b>từ khóa ở hàng 2</b> (future-proof). Nếu sau này đổi
  tên cột nhiều, sửa <code>FIELD_DEFS</code> ở đầu <code>app.js</code>.
  <b>Prompt bàn giao đầy đủ</b> nằm trong comment ở cuối file <code>index.html</code>
  (mở bằng trình soạn thảo) để mang sang chat khác làm tiếp.
</div>
`;
