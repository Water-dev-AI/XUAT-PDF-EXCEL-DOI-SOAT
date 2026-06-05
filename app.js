/* =========================================================================
   ĐỐI SOÁT ĐƠN HÀNG — app.js
   Đọc Google Sheet (API key) -> parse future-proof theo TỪ KHÓA ở hàng 2
   -> xử lý merge -> review -> xuất PDF & Excel.
   Chạy 100% trong trình duyệt. Không gửi dữ liệu đi đâu.
   ========================================================================= */

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
  const OUT_COLS = ["ngayOrder","ngayDV","orderId","maVan","sanPham","soLuong","donGiaGiam","donGiaChuaVat"];
  const OUT_LABEL = {
    ngayOrder:"Ngày Order", ngayDV:"Ngày Cung Cấp DV", orderId:"Mã Order ID",
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

  /* mã hợp lệ: order id kiểu b00..., mã vận SPXVN..., 250xxx..., chuỗi chữ-số dài.
     dùng để TÁCH phần mã ra khỏi phần text ghi chú trong cùng một ô.        */
  const looksLikeCode = tok => {
    const t = tok.trim();
    if (t.length < 6) return false;
    if (/^b0\d{6,}$/i.test(t)) return true;                 // order id
    if (/^spxvn\w+$/i.test(t)) return true;                 // shopee đơn vận
    if (/^\d{6}[a-z0-9]{6,}$/i.test(t)) return true;        // 250723UUC77M84
    if (/^[a-z0-9]{8,}$/i.test(t) && /\d/.test(t) && /[a-z]/i.test(t)) return true; // mã hỗn hợp
    if (/^cx\w+vn$/i.test(t)) return true;
    return false;
  };
  // tách 1 ô thành {codes:[...], text:"..."} — giữ thứ tự dòng
  const splitCodeText = raw => {
    if (isBlank(raw)) return {codes:[], text:""};
    const codes=[], textParts=[];
    // tách theo xuống dòng & khoảng trắng lớn & mũi tên
    const pieces = String(raw).split(/\n|-{3,}|={3,}|>+/).map(s=>s.trim()).filter(Boolean);
    pieces.forEach(piece=>{
      const toks = piece.split(/\s+/);
      const codeToks=[], txtToks=[];
      toks.forEach(tk=> looksLikeCode(tk)?codeToks.push(tk):txtToks.push(tk));
      codeToks.forEach(c=>codes.push(c));
      if(txtToks.length) textParts.push(txtToks.join(" "));
    });
    return {codes, text:textParts.join(" · ").trim()};
  };

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
        // lấy metadata: tên tab + merges + màu nền (grid)
        const metaUrl=`https://sheets.googleapis.com/v4/spreadsheets/${id}`+
          `?fields=sheets(properties(title,sheetId),merges,data(rowData(values(effectiveFormat(backgroundColor)))))`+
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

        show("info",`<span class="spin"></span>Đang tải dữ liệu ${valid.length} tab...`);
        // lấy values cho từng tab
        for(const v of valid){
          const vr=await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/`+
            encodeURIComponent(v.title)+`?valueRenderOption=FORMATTED_VALUE&key=${key}`);
          const vd=await vr.json();
          v.values=vd.values||[];
          v.merges=v.raw.merges||[];
          v.grid=v.raw.data?.[0]?.rowData||[];
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
      const warnings=[];   // ô có text lẫn mã  -> review
      const emptyDates=[]; // ngày order trống -> review
      const getCell=(ri,key)=>{const c=colMap[key];return c==null?"":(rows[ri]?.[c]??"");};

      // data bắt đầu sau header (hàng 3 trở đi = index 2). Hàng 0 = tổng, hàng 1 = tiêu đề.
      for(let ri=2; ri<rows.length; ri++){
        const rowArr=rows[ri]||[];
        const rowEmpty = rowArr.every(isBlank);

        if(rowEmpty){
          if(isGreenRow(ri) || out.length && out[out.length-1].type!=="daybreak")
            out.push({type:"daybreak", ri});
          continue;
        }

        // ---- xử lý orderId & maVan: tách mã / text ----
        const rec={type:"row", ri, raw:{}};
        OUT_COLS.concat(["mst","tenMua","diaChi","email","_donGiaTruSim","_tongSimTrang"]).forEach(k=>{
          rec.raw[k]=getCell(ri,k);
        });

        // tách code/text cho 2 cột mã
        ["orderId","maVan"].forEach(k=>{
          const {codes,text}=splitCodeText(rec.raw[k]);
          rec[k]={codes, text, original:rec.raw[k]};
          if(text){
            warnings.push({ri, field:k, fieldLabel:k==="orderId"?"Mã Order ID":"Mã Đơn Vận/Hàng",
              codes:codes.join(" · "), text, keep:codes.join(" · ")});
          }
        });

        // merge info cho cột maVan (để xác định block gộp)
        const mvCol=colMap["maVan"], oiCol=colMap["orderId"], odCol=colMap["ngayOrder"];
        rec.mergeTopMaVan = mvCol!=null && mergeAnchor[ri+"_"+mvCol] ? mergeAnchor[ri+"_"+mvCol].rows : 1;
        rec.isMergedSlave = mvCol!=null && !!mergedInto[ri+"_"+mvCol];

        // ngày order có nằm trong block merge không (thừa hưởng ngày phía trên)?
        rec.ngayOrderMerged = odCol!=null && !!mergedInto[ri+"_"+odCol];
        // hoặc: dòng con của block sản phẩm (không có mã riêng) -> nối tiếp dòng trên
        const isContinuation = rec.ngayOrderMerged ||
          (isBlank(rec.raw.orderId) && isBlank(rec.raw.maVan) && out.length &&
           out[out.length-1].type==="row");

        // ngày order trống VÀ không phải dòng nối tiếp -> mới cảnh báo
        if(isBlank(rec.raw.ngayOrder) && !isContinuation){
          emptyDates.push({ri, ngayDV:rec.raw.ngayDV, sanPham:rec.raw.sanPham, fixed:""});
          rec.emptyOrderDate=true;
        }

        // ---- tính đơn giá chưa VAT nếu thiếu ----
        rec.donGiaGiam = num(rec.raw.donGiaGiam);
        rec.soLuong = num(rec.raw.soLuong);
        let dgChua = num(rec.raw.donGiaChuaVat);
        if(dgChua==null){
          // chỉ tính khi có đơn giá có VAT hợp lệ (khác null & khác 0)
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
        const tenOk = tenStr!=="" && /\p{L}{3,}/u.test(tenStr); // có ít nhất 3 chữ cái -> tên thật
        rec.hasVatInfo = mstOk || tenOk;
        // dọn: nếu tên chỉ là số (vd "17,691,000") -> bỏ
        if(!tenOk) rec.raw.tenMua="";
        if(!mstOk) rec.raw.mst="";

        out.push(rec);
      }

      // ---- SHOPEE: gộp orderId theo block merge của maVan ----
      // Trong block merge maVan, các ô orderId con bị trống -> dồn mã/text về anchor, hiển thị merge.
      if(!isZalo){
        for(let i=0;i<out.length;i++){
          const r=out[i]; if(r.type!=="row")continue;
          if(r.mergeTopMaVan>1){
            // gom các dòng con (cùng block) — đếm theo số dòng 'row' kế tiếp thuộc merge
            let span=1, codes=[...r.orderId.codes], texts=r.orderId.text?[r.orderId.text]:[];
            for(let j=i+1;j<out.length && span<r.mergeTopMaVan;j++){
              const c=out[j]; if(c.type==="daybreak")continue;
              if(c.isMergedSlave){
                c.orderId.codes.forEach(x=>codes.push(x));
                if(c.orderId.text) texts.push(c.orderId.text);
                c.orderIdHiddenByMerge=true; span++;
              } else break;
            }
            r.orderId.codes=codes;
            r.orderId.text=texts.join(" · ");
            r.orderIdRowSpan=r.mergeTopMaVan;
            r.maVanRowSpan=r.mergeTopMaVan;
          }
        }
      }

      return {
        title:tab.title, kind:tab.kind, isZalo, month:tab.month,
        colMap, rows:out, warnings, emptyDates,
        missingChuaVat: colMap.donGiaChuaVat==null
      };
    },

    /* render + export ở phần kế tiếp (app2.js nối vào) */
    _state:()=>state,
    _util:{norm,num,fmtNum,isBlank,el,OUT_COLS,OUT_LABEL}
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

  // panel: ô có text cần quyết
  App._panelWarnings=function(sh){
    const p=el("div","review");
    const head=el("div","review-head warn-h");
    head.innerHTML=`<span>⚠ ${sh.warnings.length} ô có CHỮ GHI CHÚ lẫn trong mã — kiểm tra trước khi xuất (thuế không nên thấy)</span><span>▾</span>`;
    head.onclick=()=>p.classList.toggle("collapsed");
    const body=el("div","review-body");
    const tbl=el("table","rv");
    tbl.innerHTML="<thead><tr><th>Dòng</th><th>Cột</th><th>Mã giữ lại</th><th>Chữ ghi chú (sẽ BỎ)</th><th>Sửa mã xuất ra</th></tr></thead>";
    const tb=el("tbody");
    sh.warnings.forEach((w,i)=>{
      const tr=el("tr");
      tr.innerHTML=
        `<td>${w.ri+1}</td><td>${w.fieldLabel}</td>`+
        `<td><span class="tag code">${w.codes||'<span class=muted>—</span>'}</span></td>`+
        `<td><span class="tag txt">${w.text}</span></td>`;
      const td=el("td");
      const inp=el("input");inp.type="text";inp.value=w.keep;
      inp.oninput=e=>{w.keep=e.target.value;App._applyWarningFix(sh,w);};
      td.appendChild(inp);tr.appendChild(td);
      tb.appendChild(tr);
    });
    tbl.appendChild(tb);body.appendChild(tbl);p.appendChild(head);p.appendChild(body);
    return p;
  };

  App._applyWarningFix=function(sh,w){
    const rec=sh.rows.find(r=>r.ri===w.ri);
    if(!rec)return;
    rec[w.field].codes = w.keep.split(/[·\n,]+/).map(s=>s.trim()).filter(Boolean);
    rec[w.field].text="";
    App._refreshPreview(sh);
  };

  // panel: ngày order trống
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
        `<td style="max-width:340px;white-space:normal">${ed.sanPham||''}</td>`;
      const td=el("td");const inp=el("input");inp.type="text";inp.placeholder="vd 5/4/2025";inp.value=ed.fixed;
      inp.oninput=e=>{ed.fixed=e.target.value;const rec=sh.rows.find(r=>r.ri===ed.ri);
        if(rec)rec.raw.ngayOrder=e.target.value;App._refreshPreview(sh);};
      td.appendChild(inp);tr.appendChild(td);tb.appendChild(tr);
    });
    tbl.appendChild(tb);body.appendChild(tbl);p.appendChild(head);p.appendChild(body);
    return p;
  };

  // bảng preview giống bản in
  App._previewTable=function(sh){
    const wrap=el("div","preview-scroll");
    wrap.dataset.sheet=sh.title;
    wrap.appendChild(App._buildPreviewInner(sh));
    // legend
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
      if(r.orderIdHiddenByMerge && sh.isZalo===false) {
        // dòng con đã gộp orderId/maVan -> vẫn hiện sản phẩm, nhưng 2 cột mã bỏ (đã rowspan ở anchor)
      }
      const tr=el("tr",r.isCancel?"cancel":"");
      cols.forEach(c=>{
        // các cột mã có rowspan ở SHOPEE
        if((c==="orderId"||c==="maVan") && !sh.isZalo){
          if(r.orderIdHiddenByMerge) return; // bị anchor span phủ
          const td=el("td","code merged-cell");
          if(r.orderIdRowSpan) td.rowSpan=r.orderIdRowSpan;
          if(c==="orderId") td.innerHTML=(r.orderId.codes.join("<br>"))||
              (r.orderId.text?`<span class="muted">${r.orderId.text}</span>`:"");
          else td.innerHTML=(r.maVan.codes.join("<br>"))||"";
          tr.appendChild(td);return;
        }
        const td=el("td");
        let v="";
        if(c==="orderId") v=r.orderId? r.orderId.codes.join(", "):"";
        else if(c==="maVan") v=r.maVan? r.maVan.codes.join(", "):"";
        else if(c==="donGiaGiam"){v=fmtNum(r.donGiaGiam);td.className="num";}
        else if(c==="donGiaChuaVat"){v=fmtNum(r.donGiaChuaVat);td.className="num"+(r.computedChuaVat?" ":"");}
        else if(c==="soLuong"){v=r.soLuong==null?"":r.soLuong;td.className="num";}
        else v=r.raw[c]??"";
        td.textContent=v;
        if(c==="orderId"||c==="maVan")td.className="code";
        tr.appendChild(td);
      });
      tb.appendChild(tr);

      // dòng phụ thông tin VAT
      if(r.hasVatInfo){
        const tr2=el("tr","vatinfo");
        const parts=[];
        if(!isBlank(r.raw.mst))parts.push("MST: "+r.raw.mst);
        if(!isBlank(r.raw.tenMua))parts.push("Tên: "+r.raw.tenMua);
        if(!isBlank(r.raw.diaChi))parts.push("Đ/c: "+r.raw.diaChi);
        if(!isBlank(r.raw.email))parts.push("Email: "+r.raw.email);
        tr2.innerHTML=`<td colspan="${cols.length}">↳ ${parts.join("  |  ")}</td>`;
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
  const COLW={ngayOrder:11,ngayDV:11,orderId:20,maVan:20,sanPham:46,soLuong:5,donGiaGiam:14,donGiaChuaVat:18};

  function monthsToExport(scope){
    const st=S();
    return scope==="all"?st.months:[st.months[st.current]];
  }
  function titleFor(mo){
    // suy ra năm từ 1 ngày bất kỳ
    let yr="";
    for(const sh of mo.sheets){for(const r of sh.rows){if(r.type==="row"){
      const m=String(r.raw.ngayOrder||r.raw.ngayDV||"").match(/(20\d{2})/);if(m){yr=m[1];break;}}}if(yr)break;}
    return `Bảng đối soát đơn hàng, theo dõi bán hàng - Tháng ${mo.month}/${yr||"20XX"}`;
  }

  /* ---------- chuẩn bị ma trận dữ liệu 1 sheet để xuất ---------- */
  function sheetMatrix(sh){
    const cols=App._util.OUT_COLS_for(sh);
    const header=cols.map(c=>OUT_LABEL[c]);
    const aoa=[header];
    const merges=[];      // {s:{r,c},e:{r,c}}
    const rowMeta=[];     // 'data' | 'daybreak' | 'vatinfo'
    rowMeta.push("header");

    sh.rows.forEach(r=>{
      if(r.type==="daybreak"){aoa.push(cols.map(()=> ""));rowMeta.push("daybreak");return;}
      if(r.orderIdHiddenByMerge && !sh.isZalo){
        // dòng con của block merge: orderId & maVan để trống (sẽ bị merge phủ)
        const line=cols.map(c=>cellVal(r,c,sh,true));
        aoa.push(line);rowMeta.push(r.isCancel?"cancel":"data");
        return;
      }
      const line=cols.map(c=>cellVal(r,c,sh,false));
      aoa.push(line);rowMeta.push(r.isCancel?"cancel":"data");

      // merge cho block anchor (SHOPEE)
      if(r.orderIdRowSpan && r.orderIdRowSpan>1 && !sh.isZalo){
        const top=aoa.length-1;
        ["orderId","maVan"].forEach(k=>{
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
        const vrow=cols.map(()=> "");vrow[0]="↳ "+parts.join("   |   ");
        aoa.push(vrow);rowMeta.push("vatinfo");
        merges.push({s:{r:aoa.length-1,c:0},e:{r:aoa.length-1,c:cols.length-1}});
      }
    });
    return {cols,aoa,merges,rowMeta};
  }
  function cellVal(r,c,sh,slave){
    if((c==="orderId"||c==="maVan")){
      if(slave) return "";
      return c==="orderId"?r.orderId.codes.join("\n"):r.maVan.codes.join("\n");
    }
    if(c==="donGiaGiam")return r.donGiaGiam==null?"":r.donGiaGiam;
    if(c==="donGiaChuaVat")return r.donGiaChuaVat==null?"":Math.round(r.donGiaChuaVat);
    if(c==="soLuong")return r.soLuong==null?"":r.soLuong;
    return r.raw[c]??"";
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
        const numCols=new Set([cols.indexOf("soLuong"),cols.indexOf("donGiaGiam"),cols.indexOf("donGiaChuaVat")]);
        const range={s:{r:0,c:0},e:{r:data.length-1,c:cols.length-1}};
        for(let r=0;r<data.length;r++){
          for(let c=0;c<cols.length;c++){
            const addr=XLSX.utils.encode_cell({r,c});
            let v=data[r][c];
            const m=meta[r];
            const cell={};
            // kiểu số thật cho các cột số (trừ dòng title/header/vatinfo)
            if(numCols.has(c) && m!=="title" && m!=="header" && m!=="vatinfo" && v!=="" && v!=null && !isNaN(Number(v))){
              cell.t="n"; cell.v=Number(v); cell.z=numFmt;
            } else {
              cell.t="s"; cell.v=(v==null?"":String(v));
            }
            // style
            const st={border:BORDER,alignment:{vertical:"top",wrapText:true}};
            if(m==="title"){st.font={bold:true,sz:13};st.alignment={horizontal:"center",vertical:"center"};delete st.border;}
            else if(m==="header"){st.font={bold:true,color:{rgb:"FFFFFF"},sz:9};
              st.fill={fgColor:{rgb:"10243F"}};st.alignment={horizontal:"center",vertical:"center",wrapText:true};}
            else if(m==="daybreak"){st.fill={fgColor:{rgb:"CFE8D8"}};}
            else if(m==="vatinfo"){st.fill={fgColor:{rgb:"FFF7E6"}};st.font={italic:true,sz:8,color:{rgb:"5A4200"}};}
            else if(m==="cancel"){st.fill={fgColor:{rgb:"FDECEA"}};}
            if(numCols.has(c)&&m!=="title"&&m!=="header") st.alignment={...st.alignment,horizontal:"right"};
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

  /* ---------- PDF (qua cửa sổ in của trình duyệt: ngang, font hẹp) ---------- */
  App.exportPDF=function(scope){
    const mos=monthsToExport(scope);
    let html=`<html><head><meta charset="utf-8"><title>Đối soát</title><style>
      @page{size:A4 landscape;margin:8mm}
      *{box-sizing:border-box}
      html,body{margin:0;padding:0}
      body{font-family:"Times New Roman",serif;color:#111}
      .doc-title{text-align:center;font-size:14pt;font-weight:bold;margin:0 0 8px}
      .sheet-title{font-size:10.5pt;font-weight:bold;margin:8px 0 4px}
      table{border-collapse:collapse;width:100%;table-layout:fixed}
      th,td{border:0.5pt solid #555;padding:2px 3px;font-size:7pt;word-wrap:break-word;
        overflow-wrap:break-word;vertical-align:top;line-height:1.2}
      th{background:#10243f !important;color:#fff !important;font-size:6.8pt;text-align:center;
        -webkit-print-color-adjust:exact;print-color-adjust:exact}
      td.num{text-align:right;font-variant-numeric:tabular-nums}
      td.code{font-family:"Consolas",monospace;font-size:6.4pt;word-break:break-all}
      tr.daybreak td{background:#cfe8d8 !important;height:3px;padding:0;
        -webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr.cancel td{background:#fdecea !important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr.vatinfo td{background:#fff7e6 !important;font-size:6.2pt;font-style:italic;color:#5a4200;
        text-align:left;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      tr{page-break-inside:avoid}
      thead{display:table-header-group}
      .sheet-section{page-break-inside:auto}
      .sheet-section+.sheet-section{page-break-before:always}
      .month-section+.month-section{page-break-before:always}
    </style></head><body>`;

    mos.forEach((mo,mi)=>{
      html+=`<div class="month-section">`;
      mo.sheets.forEach((sh,si)=>{
        const {cols,aoa,rowMeta,merges}=sheetMatrix(sh);
        html+=`<div class="sheet-section">`;
        if(si===0) html+=`<div class="doc-title">${titleFor(mo)}</div>`;
        html+=`<div class="sheet-title">${sh.title} — ${sh.kind}</div>`;
        // map rowspan từ merges. merges.r tính trên aoa (gồm dòng header ở index 0)
        const spanAt={}; const skip={};
        merges.forEach(m=>{
          spanAt[m.s.r+"_"+m.s.c]=m.e.r-m.s.r+1;
          for(let rr=m.s.r+1;rr<=m.e.r;rr++) skip[rr+"_"+m.s.c]=1;
        });
        const totalW=cols.reduce((a,c)=>a+(COLW[c]||12),0);
        html+="<table><colgroup>"+cols.map(c=>`<col style="width:${((COLW[c]||12)/totalW*100).toFixed(2)}%">`).join("")+"</colgroup>";
        html+="<thead><tr>"+cols.map(c=>`<th>${OUT_LABEL[c]}</th>`).join("")+"</tr></thead><tbody>";
        for(let ri=0;ri<aoa.length;ri++){
          const meta=rowMeta[ri];
          if(meta==="header")continue;
          if(meta==="daybreak"){html+=`<tr class="daybreak"><td colspan="${cols.length}"></td></tr>`;continue;}
          if(meta==="vatinfo"){html+=`<tr class="vatinfo"><td colspan="${cols.length}">${esc(aoa[ri][0])}</td></tr>`;continue;}
          html+=`<tr class="${meta==='cancel'?'cancel':''}">`;
          for(let ci=0;ci<cols.length;ci++){
            if(skip[ri+"_"+ci])continue;
            const c=cols[ci];const isNum=(c==="donGiaGiam"||c==="donGiaChuaVat"||c==="soLuong");
            const isCode=(c==="orderId"||c==="maVan");
            let v=aoa[ri][ci];
            if(isNum&&v!==""&&v!=null)v=Number(v).toLocaleString("vi-VN");
            v=esc(v).replace(/\n/g,"<br>");
            const sp=spanAt[ri+"_"+ci]?` rowspan="${spanAt[ri+"_"+ci]}"`:"";
            html+=`<td class="${isNum?'num':''}${isCode?' code':''}"${sp}>${v}</td>`;
          }
          html+="</tr>";
        }
        html+="</tbody></table></div>";
      });
      html+="</div>";
    });
    html+="</body></html>";

    const w=window.open("","_print");
    w.document.write(html);w.document.close();
    w.onload=()=>{setTimeout(()=>w.print(),300);};
  };

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
        ngayOrder:find("ngày","order")>=0?find("ngày","order"):0,
        ngayDV:find("cung","cấp"), orderId:find("order","id"),
        maVan:isZalo?-1:find("đơn","vận"),
        sanPham:find("dịch","vụ")>=0?find("dịch","vụ"):find("sản","phẩm"),
        soLuong:find("sl")>=0?find("sl"):find("số","lượng"),
        donGiaGiam:find("sau","giảm")>=0&&header.findIndex(h=>/sau giảm/i.test(h)&&!/chưa vat/i.test(h)),
        donGiaChuaVat:find("chưa","vat")
      };
      // build merges map từ ws['!merges'] cho cột mã
      const mergeRows={};
      (ws["!merges"]||[]).forEach(m=>{ if(m.e.r>m.s.r) mergeRows[m.s.r]=Math.max(mergeRows[m.s.r]||1,m.e.r-m.s.r+1); });

      const rows=[];
      for(let r=2;r<aoa.length;r++){
        const row=aoa[r]; if(!row)continue;
        const allEmpty=row.every(v=>String(v).trim()==="");
        // dòng phụ VAT: bắt đầu bằng ↳
        if(String(row[0]).trim().startsWith("↳")){
          const prev=rows[rows.length-1];
          if(prev){prev.hasVatInfo=true;prev._vatLine=String(row[0]).replace(/^↳\s*/,"");}
          continue;
        }
        if(allEmpty){rows.push({type:"daybreak",ri:r});continue;}
        const g=k=>{const i=ci[k];return (i==null||i<0)?"":(row[i]??"");};
        const rec={type:"row",ri:r,raw:{
          ngayOrder:g("ngayOrder"),ngayDV:g("ngayDV"),sanPham:g("sanPham"),
          mst:"",tenMua:"",diaChi:"",email:""},
          orderId:{codes:String(g("orderId")).split(/\n/).filter(Boolean),text:""},
          maVan:{codes:String(g("maVan")).split(/\n/).filter(Boolean),text:""},
          soLuong:num(g("soLuong")),donGiaGiam:num(g("donGiaGiam")),
          donGiaChuaVat:num(g("donGiaChuaVat"))
        };
        rec.isCancel=(rec.soLuong!=null&&rec.soLuong<0);
        if(mergeRows[r]){rec.orderIdRowSpan=mergeRows[r];rec.maVanRowSpan=mergeRows[r];}
        rows.push(rec);
      }
      const sh={title:name,kind,isZalo,month,rows,warnings:[],emptyDates:[],
        missingChuaVat:false,colMap:{}};
      (byMonth[month]=byMonth[month]||[]).push(sh);
    });
    const months=[];
    Object.keys(byMonth).map(Number).sort((a,b)=>a-b).forEach(m=>months.push({month:m,sheets:byMonth[m]}));
    return months;
  };
})();
