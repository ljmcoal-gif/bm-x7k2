const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

// ---------- SK palette (deck spec) ----------
const BG   = "FBF9F5";
const INK  = "1B1714";
const RED  = "C1121F";
const POS  = "1E7A4D";
const CMP  = "D8CFC2";   // 비교계열
const CMPD = "B9AE9C";
const LINE = "E5DECF";
const LIND = "CFC6B4";
const MUTE = "8C8473";
const NAVY = "2A3A5C";   // NAV 마커
const P1   = "1B1714";   // plant1 = ink
const P2   = "C1121F";   // plant2 = red
const WHITE= "FFFFFF";

const HEAD = "Bookman Old Style";   // serif header (safe-list)
const BODY = "Calibri";             // sans body (safe-list)
const MONO = "Courier New";         // numbers

async function icon(Comp, color = INK, size = 256){
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#"+color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name:"SK", width:13.333, height:9.43 });  // ~1920x1358 ratio
  pres.layout = "SK";
  pres.author = "SK · CSO"; pres.title = "US LFP/ESS Cell Platform — Valuation & Financing";
  const W = 13.333, Hh = 9.43, MX = 0.7;

  // pre-render icons
  const ic = {
    shield:  await icon(fa.FaShieldAlt, RED),
    bolt:    await icon(fa.FaBolt, RED),
    micro:   await icon(fa.FaMicrochip, RED),
    flag:    await icon(fa.FaFlagUsa, RED),
    gov:     await icon(fa.FaLandmark, RED),
    handsh:  await icon(fa.FaHandshake, RED),
    coins:   await icon(fa.FaCoins, RED),
    lock:    await icon(fa.FaLock, RED),
    chart:   await icon(fa.FaChartLine, RED),
    layers:  await icon(fa.FaLayerGroup, RED),
    industry:await icon(fa.FaIndustry, INK),
    battery: await icon(fa.FaBatteryThreeQuarters, RED),
  };

  // ---------- shared chrome helpers ----------
  function base(slide){
    slide.background = { color: BG };
  }
  function header(slide, eyebrow, pageNo, titleRuns, dek){
    slide.addText([
      { text:"SK", options:{ fontFace:HEAD, bold:true, color:INK, fontSize:20, charSpacing:1 } },
      { text:"   ", options:{} },
      { text:eyebrow, options:{ fontFace:BODY, bold:true, color:RED, fontSize:12, charSpacing:2 } },
    ], { x:MX, y:0.5, w:10.5, h:0.4, margin:0, valign:"middle" });
    slide.addText(pageNo, { x:W-2.3, y:0.5, w:1.6, h:0.4, margin:0, align:"right",
      fontFace:MONO, color:MUTE, fontSize:12, valign:"middle" });
    slide.addText(titleRuns, { x:MX, y:0.92, w:W-2*MX, h:0.85, margin:0,
      fontFace:HEAD, bold:true, color:INK, fontSize:30, valign:"middle" });
    if(dek) slide.addText(dek, { x:MX, y:1.78, w:W-2*MX, h:0.55, margin:0,
      fontFace:BODY, color:MUTE, fontSize:14.5, valign:"middle" });
    // thin rule under header (whitespace separation, full-width hairline is fine as a rule not a stripe)
    slide.addShape(pres.shapes.LINE, { x:MX, y:2.42, w:W-2*MX, h:0, line:{ color:INK, width:1.4 } });
  }
  function footer(slide){
    slide.addShape(pres.shapes.LINE, { x:MX, y:Hh-0.62, w:W-2*MX, h:0, line:{ color:LIND, width:0.75 } });
    slide.addText([
      { text:"Project [ ● ]", options:{ bold:true, color:INK } },
      { text:"  ·  Strictly Private & Confidential", options:{ color:MUTE } },
    ], { x:MX, y:Hh-0.55, w:6, h:0.32, margin:0, fontFace:BODY, fontSize:10.5, valign:"middle" });
    slide.addText("DRAFT — FOR DISCUSSION ONLY", { x:W/2-2.5, y:Hh-0.55, w:5, h:0.32, margin:0,
      align:"center", fontFace:BODY, bold:true, color:RED, fontSize:10.5, charSpacing:2, valign:"middle" });
    slide.addText("2026. 6", { x:W-2.2, y:Hh-0.55, w:1.5, h:0.32, margin:0, align:"right",
      fontFace:MONO, color:MUTE, fontSize:10.5, valign:"middle" });
  }
  const sh = () => ({ type:"outer", color:"000000", blur:7, offset:2, angle:90, opacity:0.10 });

  // ============================================================
  // SLIDE 1 — FOOTBALL FIELD
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "VALUATION SUMMARY · FOOTBALL FIELD", "01 / 12",
      [{ text:"통합 플랫폼 기업가치 — 3개 방법론 수렴", options:{} }],
      "DCF(DC블록·9%·NOL $1.5bn 반영) $3.5bn · EV/EBITDA ex-AMPC 8–14x $2.8–4.9bn · NAV(SOTP) $5.0bn · 결론 EV ≈ $3.5bn");

    // chart geometry
    const cx = MX, cy = 2.95, chartW = 8.55, chartH = 4.55;
    const axX = cx + 2.55;            // bars start
    const axW = chartW - 2.55 - 0.15;
    const v2x = v => axX + (v/7.0)*axW;   // value -> x  (0~7bn scale)

    // gridlines
    const ticks = [0,1,2,3,4,5,6,7];
    s.addText("Enterprise Value (USD bn)", { x:cx, y:cy-0.32, w:5, h:0.3, margin:0,
      fontFace:BODY, color:MUTE, fontSize:11, bold:true });
    const gTop = cy+0.1, gBot = cy+chartH-0.55;
    // concluded band 2.8 - 3.8 (DCF 민감도 ~ ex-AMPC 11x 부근)
    s.addShape(pres.shapes.RECTANGLE, { x:v2x(3.0), y:gTop, w:v2x(4.0)-v2x(3.0), h:gBot-gTop,
      fill:{ color:RED, transparency:92 }, line:{ color:RED, width:0.75, transparency:60 } });
    ticks.forEach(t=>{
      s.addShape(pres.shapes.LINE, { x:v2x(t), y:gTop, w:0, h:gBot-gTop,
        line:{ color: (t===0||t===7)?LIND:LINE, width:(t===0||t===7)?1:0.75 } });
      s.addText(t.toFixed(1), { x:v2x(t)-0.4, y:gBot+0.05, w:0.8, h:0.3, margin:0,
        align:"center", fontFace:MONO, color:MUTE, fontSize:11 });
    });
    // concluded dashed line @3.45 (DCF EV base, NOL 반영) + tag
    s.addShape(pres.shapes.LINE, { x:v2x(3.45), y:gTop, w:0, h:gBot-gTop, line:{ color:RED, width:1.5, dashType:"dash" } });
    s.addText("결론 EV ≈ $3.5bn", { x:v2x(3.45)+0.06, y:gTop-0.02, w:2.3, h:0.32, margin:0.03,
      fill:{ color:RED }, color:WHITE, fontFace:BODY, bold:true, fontSize:11, align:"center", valign:"middle" });

    // rows  (엑셀 실제값: DC블록 $155 · ex-AMPC EBITDA 9% · NOL 반영)
    const rows = [
      ["DCF (DC블록·9%·NOL 반영)","WACC 9~12%·TGR 2%·AMPC $45·NOL $1.5bn", 2.97, 4.11, true],
      ["EV/EBITDA ex-AMPC (2029E)","fwd 8–14x · EBITDA $350mm", 2.80, 4.90, false],
      ["EV/EBITDA ex-AMPC (run-rate)","fwd 8–14x · EBITDA $441mm", 3.53, 6.17, false],
      ["순자산 NAV (SOTP)","SKBA $2.0 + SKOT $3.0bn", 5.00, 5.00, false],
    ];
    const rowH = (gBot-gTop)/rows.length;
    rows.forEach((r,i)=>{
      const ry = gTop + i*rowH + rowH/2;
      s.addText([{text:r[0],options:{bold:true,fontSize:13,color:INK,breakLine:true}},
                 {text:r[1],options:{fontSize:9.5,color:MUTE}}],
        { x:cx, y:ry-0.42, w:2.45, h:0.84, margin:0, fontFace:BODY, valign:"middle" });
      const bx=v2x(r[2]), bw=v2x(r[3])-v2x(r[2]);
      const isPoint = Math.abs(r[3]-r[2])<0.01;
      if(isPoint){
        // single-point estimate (NAV) → diamond marker
        s.addShape(pres.shapes.DIAMOND, { x:bx-0.13, y:ry-0.16, w:0.32, h:0.32,
          fill:{ color:NAVY }, line:{ color:NAVY, width:1 } });
        s.addText("$"+r[2].toFixed(1)+"bn", { x:bx+0.24, y:ry-0.16, w:1.2, h:0.32, margin:0, align:"left",
          fontFace:MONO, fontSize:10.5, bold:true, color:NAVY, valign:"middle" });
      } else {
        s.addShape(pres.shapes.RECTANGLE, { x:bx, y:ry-0.16, w:bw, h:0.32,
          fill:{ color: r[4]?INK:CMP }, line:{ color: r[4]?INK:CMPD, width:1 } });
        s.addText(r[2].toFixed(2), { x:bx-0.62, y:ry-0.16, w:0.55, h:0.32, margin:0, align:"right",
          fontFace:MONO, fontSize:10.5, bold:true, color:INK, valign:"middle" });
        s.addText(r[3].toFixed(2), { x:bx+bw+0.07, y:ry-0.16, w:0.55, h:0.32, margin:0, align:"left",
          fontFace:MONO, fontSize:10.5, bold:true, color:INK, valign:"middle" });
      }
    });

    // legend
    const ly = cy+chartH+0.05;
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:ly+0.04, w:0.28, h:0.16, fill:{color:INK}, line:{color:INK,width:1} });
    s.addText("DCF (key)", { x:cx+0.34, y:ly-0.04, w:1.4, h:0.3, margin:0, fontFace:BODY, fontSize:10.5, valign:"middle" });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+1.7, y:ly+0.04, w:0.28, h:0.16, fill:{color:CMP}, line:{color:CMPD,width:1} });
    s.addText("방법론 레인지", { x:cx+2.04, y:ly-0.04, w:1.6, h:0.3, margin:0, fontFace:BODY, fontSize:10.5, valign:"middle" });
    s.addShape(pres.shapes.RECTANGLE, { x:cx+3.75, y:ly+0.04, w:0.28, h:0.16, fill:{color:RED,transparency:80}, line:{color:RED,width:1} });
    s.addText("결론 밴드 $3.0–4.0bn", { x:cx+4.09, y:ly-0.04, w:2.6, h:0.3, margin:0, fontFace:BODY, fontSize:10.5, valign:"middle" });

    // NAV side cards
    const nx = MX + chartW + 0.55, nw = W - nx - MX;
    s.addText("순자산가치 — 엔티티별 (SOTP)", { x:nx, y:cy-0.32, w:nw, h:0.35, margin:0,
      fontFace:HEAD, bold:true, fontSize:14, color:INK });
    s.addShape(pres.shapes.LINE, { x:nx, y:cy+0.06, w:nw, h:0, line:{ color:INK, width:1.2 } });
    function navCard(y, title, tag, rows, edgeColor){
      const ch = rows.length*0.34 + 0.62;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:nx, y, w:nw, h:ch, rectRadius:0.06,
        fill:{ color:WHITE }, line:{ color:LIND, width:1 }, shadow:sh() });
      s.addImage({ data: ic.industry, x:nx+0.16, y:y+0.16, w:0.26, h:0.26 });
      s.addText(title, { x:nx+0.5, y:y+0.12, w:nw-1.45, h:0.35, margin:0, fontFace:BODY, bold:true, fontSize:13.5, color:INK, valign:"middle" });
      s.addText(tag, { x:nx+nw-1.35, y:y+0.12, w:1.2, h:0.35, margin:0, align:"right", fontFace:BODY, fontSize:10, color:MUTE, valign:"middle" });
      rows.forEach((r,i)=>{
        const ry = y+0.56+i*0.34;
        s.addText(r[0], { x:nx+0.18, y:ry, w:nw*0.55, h:0.3, margin:0, fontFace:BODY, fontSize:11, color:MUTE, valign:"middle" });
        s.addText(r[1], { x:nx+nw*0.45-0.18, y:ry, w:nw*0.55, h:0.3, margin:0, align:"right",
          fontFace:MONO, fontSize:11.5, bold:true, color: r[2]?RED:INK, valign:"middle" });
        if(i<rows.length-1) s.addShape(pres.shapes.LINE, { x:nx+0.18, y:ry+0.32, w:nw-0.36, h:0, line:{ color:LINE, width:0.5, dashType:"dash" } });
      });
      return ch;
    }
    let yy = cy+0.2;
    const h1 = navCard(yy, "공장 1 — SKBA (조지아)", "가동중",
      [["순자산 (NAV)","~$2.0 bn ★",true],["캐파 (파우치)","21.2 GWh",false],["개조 capex","$0.4 bn",false],["가동(COD)","1-1동 '27 / 1-2동 '28",false]], P1);
    yy += h1 + 0.22;
    const h2 = navCard(yy, "공장 2 — SKOT (테네시)", "'28 가동",
      [["순자산 (NAV)","$3.0 bn",true],["캐파 (각형+파우치)","29.7 GWh",false],["개조 capex","$1.0 bn",false],["DOE loan","$4.0 bn",false]], P2);
    yy += h2 + 0.16;
    s.addText("SKBA NAV는 추정(SK온 출자 ~4.8조−누적결손, DART 확인要) · SKOT는 분할 PPA. 합산 NAV ~$5.0bn vs DCF EV ~$3.5bn.",
      { x:nx, y:yy, w:nw, h:0.5, margin:0, fontFace:BODY, fontSize:9.5, color:MUTE, italic:true });

    // source note
    s.addText([
      { text:"제약 반영:  ", options:{ bold:true, color:INK } },
      { text:"공장 2의 $4.0bn DOE loan → 추가 federal funding 불가 + 모회사(공장1) 외부 지분 20% 초과 이전 시 CoC 조항 발동. 양 제약이 신규 equity 여력을 제한해 통합 EV에 디스카운트로 작용.", options:{ color:MUTE } },
    ], { x:MX, y:Hh-1.18, w:W-2*MX, h:0.5, margin:0, fontFace:BODY, fontSize:10.5, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 2 — PEER MULTIPLES
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "COMPARABLE COMPANIES · TRADING MULTIPLES", "02 / 12",
      [{ text:"배터리 셀 & ESS Peer Group — 트레이딩 멀티플", options:{} }],
      "상장 셀 제조사 & ESS 시스템 peer · 시장데이터 2026. 6 기준 · 수치 FY/CY2025A–2026E, indicative");

    const cols = ["Company","Mkt Cap","Ent. Value","Rev '26E","Rev Gr.","EBITDA mgn","EV/Rev '26E","EV/EBITDA '26E","P/B"];
    const colW = [3.05, 1.15, 1.15, 1.05, 0.92, 1.18, 1.2, 1.35, 0.88];
    // header row
    const hdr = cols.map((c,i)=>({ text:c, options:{
      align: i===0?"left":"right", bold:true, color: i>=6?RED:MUTE, fontSize:10.5,
      fontFace:BODY, fill:{color:BG}, valign:"bottom",
      border:[{type:"none"},{type:"none"},{pt:1.5,color: i>=6?RED:INK},{type:"none"}] } }));
    const rows=[hdr];
    const seg = label => rows.push([{ text:label, options:{ colspan:9, bold:true, color:RED, fontSize:10.5,
      fontFace:BODY, fill:{color:BG}, charSpacing:1, border:[{type:"none"},{type:"none"},{type:"none"},{type:"none"}] } }]);
    const dataRow = (co, sub, vals, opts={}) => {
      const r=[{ text:[{text:co+"\n",options:{bold:true,fontSize:12.5,color:INK}},{text:sub,options:{fontSize:9,color:MUTE}}],
        options:{ align:"left", fontFace:BODY, valign:"middle", fill:{color:opts.fill||BG},
        border:[{type:"none"},{type:"none"},{pt:0.5,color:LINE},{type:"none"}] } }];
      vals.forEach((v,i)=>{
        const isMlt = i>=5;
        r.push({ text:v.t, options:{ align:"right", fontFace: v.mono===false?BODY:MONO,
          bold: isMlt||v.bold, italic:v.it, color: v.pos?POS : (v.nm?MUTE:INK),
          fontSize: isMlt?13:11.5, fill:{color:opts.fill||BG}, valign:"middle",
          border:[{type:"none"},{type:"none"},{pt:0.5,color:LINE},{type:"none"}] } });
      });
      rows.push(r);
    };

    seg("— 대형 셀 제조사 —");
    dataRow("LG Energy Solution","KS:373220 · EV/ESS 셀",
      [{t:"₩105.0T"},{t:"₩118T"},{t:"₩31T"},{t:"+9%",pos:true},{t:"~11%"},{t:"3.8x"},{t:"11x"},{t:"2.4x"}]);
    dataRow("Samsung SDI","KS:006400 · EV/ESS 셀",
      [{t:"₩28.0T"},{t:"₩36T"},{t:"₩23T"},{t:"+6%",pos:true},{t:"~9%"},{t:"1.6x"},{t:"8.6x"},{t:"1.1x"}]);
    dataRow("CATL","SZ:300750 · EV/ESS 셀",
      [{t:"¥1,520B"},{t:"¥1,380B"},{t:"¥460B"},{t:"+18%",pos:true},{t:"23%"},{t:"3.0x"},{t:"13x"},{t:"4.6x"}]);
    seg("— 미국 / ESS 시스템 · 첨단셀 Peer —");
    dataRow("Fluence Energy","NASDAQ:FLNC · US ESS integrator",
      [{t:"$4.3B"},{t:"$4.1B"},{t:"$3.4B"},{t:"+50%",pos:true},{t:"~1.5%"},{t:"1.2x"},{t:"37x*"},{t:"5.1x"}]);
    dataRow("Tesla — Energy 부문","NASDAQ:TSLA · Megapack LFP",
      [{t:"n/a (seg.)"},{t:"n/a"},{t:"$14B"},{t:"+40%",pos:true},{t:"~28%"},{t:"≈2.0x"},{t:"≈8x"},{t:"—"}]);
    dataRow("Amprius Technologies","NYSE:AMPX · 첨단셀",
      [{t:"$0.9B"},{t:"$0.8B"},{t:"$0.20B"},{t:"+150%",pos:true},{t:"neg."},{t:"4.0x"},{t:"n.m.",nm:true,it:true},{t:"3.2x"}]);
    dataRow("EVE Energy","SZ:300014 · ESS-heavy",
      [{t:"¥130B"},{t:"¥150B"},{t:"¥75B"},{t:"+35%",pos:true},{t:"15%"},{t:"2.0x"},{t:"13x"},{t:"3.0x"}]);
    // median
    const med=[{ text:"PEER MEDIAN", options:{ align:"left", bold:true, color:RED, fontSize:11.5, fontFace:BODY,
      charSpacing:1, fill:{color:"F2ECE0"}, valign:"middle", border:[{pt:1.5,color:INK},{type:"none"},{pt:1.5,color:INK},{type:"none"}] } }];
    [["—"],["—"],["—"],["+18%"],["~11%"],["2.0x"],["8–13x"],["3.0x"]].forEach((v,i)=>{
      med.push({ text:v[0], options:{ align:"right", bold:true, fontFace: i<3?BODY:MONO, color:INK, fontSize: i>=5?13:11.5,
        fill:{color:"F2ECE0"}, valign:"middle", border:[{pt:1.5,color:INK},{type:"none"},{pt:1.5,color:INK},{type:"none"}] } });
    });
    rows.push(med);

    s.addTable(rows, { x:MX, y:2.55, w:W-2*MX, colW, rowH:0.40, valign:"middle", margin:[2,4,2,4] });

    s.addText([
      { text:"타겟 read-through:  ", options:{ bold:true, color:INK } },
      { text:"US LFP/ESS DC블록 · AMPC 적격 · ex-AMPC EBITDA 9% 가정. peer forward EV/EBITDA 8–14x를 우리 ", options:{ color:INK } },
      { text:"ex-AMPC EBITDA", options:{ bold:true, color:RED } },
      { text:"에 적용(AMPC는 2033 소멸 한시 보조금이라 영구 자본화 부적절). 단 peer EBITDA에도 AMPC 포함 → 보수적 비교.", options:{ color:INK } },
    ], { x:MX, y:Hh-2.25, w:W-2*MX, h:0.5, margin:0, fontFace:BODY, fontSize:10, valign:"top" });

    // implied valuation mini-block
    const ivy=Hh-1.65;
    s.addShape(pres.shapes.RECTANGLE, { x:MX, y:ivy, w:W-2*MX, h:0.4, fill:{color:INK} });
    s.addText("우리 회사 적용 — EV/EBITDA ex-AMPC (forward)", { x:MX+0.15, y:ivy, w:5.5, h:0.4, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:10.5, valign:"middle" });
    s.addText("Fwd EBITDA ex-AMPC $350M (2029E) / run-rate $441M  ·  indicative, CapIQ·Daloopa refresh 要", { x:MX+5.6, y:ivy, w:5.5, h:0.4, margin:0, fontFace:BODY, color:"C9C1B2", fontSize:8.5, valign:"middle", align:"right" });
    const ivd=[["8x (보수)","$2.8bn"],["11x (기준)","$3.9bn"],["14x (강세)","$4.9bn"],["DCF 대조","$3.5bn"]];
    const ivw=(W-2*MX)/4;
    ivd.forEach((d,i)=>{
      const ix=MX+i*ivw;
      s.addShape(pres.shapes.RECTANGLE, { x:ix, y:ivy+0.4, w:ivw, h:0.46, fill:{color: i===3?"F2ECE0":BG}, line:{color:LINE,width:0.5} });
      s.addText([{text:d[0]+"  ",options:{fontSize:9.5,color:MUTE}},{text:d[1],options:{bold:true,fontSize:13,color: i===3?INK:RED,fontFace:MONO}}],
        { x:ix+0.1, y:ivy+0.4, w:ivw-0.2, h:0.46, margin:0, fontFace:BODY, align:"center", valign:"middle" });
    });
    footer(s);
  }

  // ============================================================
  // SLIDE 3 — DCF & PROJECTION ASSUMPTIONS
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "DCF & PROJECTION ASSUMPTIONS", "03 / 12",
      [{ text:"주요 DCF · 프로젝션 가정", options:{} }],
      "Base case 전제 · 단계적 COD(SKBA 1-1동 '27 / 1-2동·SKOT '28) · AMPC $45 · ex-AMPC EBITDA 9% · ★ = 미검증 가정");

    const colX = [MX, W/2+0.15];
    const colW2 = (W-2*MX-0.3)/2;

    function sectionTitle(x,y,n,title){
      s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.34, h:0.34, fill:{color:RED}, line:{color:RED,width:1} });
      s.addText(String(n), { x, y, w:0.34, h:0.34, margin:0, align:"center", valign:"middle", fontFace:BODY, bold:true, color:WHITE, fontSize:15 });
      s.addText(title, { x:x+0.5, y:y-0.03, w:colW2-0.5, h:0.4, margin:0, fontFace:HEAD, bold:true, fontSize:17, color:INK, valign:"middle" });
      s.addShape(pres.shapes.LINE, { x, y:y+0.46, w:colW2, h:0, line:{ color:INK, width:1.2 } });
    }
    function kv(x,y,k,sub,v,ph){
      s.addText([{text:k,options:{fontSize:12,color:INK,bold:false,breakLine:!!sub}},
                 ...(sub?[{text:sub,options:{fontSize:8.5,color:MUTE}}]:[])],
        { x, y, w:colW2*0.62, h:0.42, margin:0, fontFace:BODY, valign:"middle" });
      s.addText(v, { x:x+colW2*0.45, y, w:colW2*0.55, h:0.42, margin:0, align:"right",
        fontFace:MONO, bold:true, fontSize:12, color: ph?RED:INK, valign:"middle" });
      s.addShape(pres.shapes.LINE, { x, y:y+0.42, w:colW2, h:0, line:{ color:LINE, width:0.5 } });
    }

    // LEFT col 1: operations
    let lx=colX[0], ly=2.7;
    sectionTitle(lx,ly,1,"운영 · 매출 가정"); ly+=0.62;
    const opRows=[
      ["제품 / 시장","AMPC 적격 US 생산","LFP DC블록 · ESS",false],
      ["DC블록 ASP","셀+모듈+컨테이너","$155 /kWh",false],
      ["ex-AMPC EBITDA","메자닌 $1bn 적격 수준","9.0%",false],
      ["AMPC 단가","45X 셀+모듈","$45 → 0 ('33)",false],
      ["명판용량 — SKBA","1-1동 '27 + 1-2동 '28","21.2 GWh",false],
      ["명판용량 — SKOT","각형+파우치 · '28","29.7 GWh",false],
    ];
    opRows.forEach(r=>{ kv(lx,ly,r[0],r[1],r[2],r[3]); ly+=0.42; });
    // capex sub (tinted)
    s.addShape(pres.shapes.RECTANGLE, { x:lx, y:ly, w:colW2, h:0.42, fill:{color:"EFE9DC"}, line:{color:"EFE9DC",width:1} });
    s.addText("개조 capex (총)", { x:lx+0.1, y:ly, w:colW2*0.6, h:0.42, margin:0, fontFace:BODY, bold:true, fontSize:12, color:INK, valign:"middle" });
    s.addText("$1.4 bn", { x:lx+colW2*0.45, y:ly, w:colW2*0.53, h:0.42, margin:0, align:"right", fontFace:MONO, bold:true, fontSize:12, color:INK, valign:"middle" });
    ly+=0.42;
    kv(lx,ly,"└ 공장1 / 공장2","","$0.4 / $1.0 bn",false); ly+=0.52;

    // section 2: COD timeline
    sectionTitle(lx,ly,2,"단계적 가동 (COD) 타임라인"); ly+=0.62;
    const tlY=ly+0.35, tlX=lx+0.1, tlW=colW2-0.2;
    s.addShape(pres.shapes.LINE, { x:tlX, y:tlY, w:tlW, h:0, line:{ color:LIND, width:1.4 } });
    const yrs=[["2026","착공·개조","idle"],["2027","SKBA 1-1동\n가동","d1"],["2028","SKBA 1-2동\n+ SKOT","d2"],
               ["2029","ramp-up","idle"],["2030","full\nrun-rate","idle"],["2031+","AMPC\nphase-out","idle"]];
    const stepW=tlW/yrs.length;
    yrs.forEach((y,i)=>{
      const px=tlX+stepW*i+stepW/2;
      s.addText(y[0], { x:px-0.55, y:tlY-0.42, w:1.1, h:0.3, margin:0, align:"center", fontFace:MONO, bold:true, fontSize:12, color:INK });
      const dotC = y[2]==="d1"?P1 : y[2]==="d2"?P2 : WHITE;
      s.addShape(pres.shapes.OVAL, { x:px-0.1, y:tlY-0.1, w:0.2, h:0.2, fill:{color:dotC}, line:{color: y[2]==="idle"?CMPD:dotC, width:2} });
      s.addText(y[1], { x:px-0.7, y:tlY+0.16, w:1.4, h:0.5, margin:0, align:"center", fontFace:BODY, fontSize:9.5,
        bold: y[2]!=="idle", color: y[2]==="idle"?MUTE:INK });
    });
    ly = tlY+0.9;
    // (constraint callout moved to bottom full-width note)

    // RIGHT col: DCF params + sensitivity
    let rx=colX[1], ry=2.7;
    sectionTitle(rx,ry,3,"DCF 핵심 파라미터"); ry+=0.62;
    const dcfRows=[
      ["예측기간","explicit forecast","2026 – 2035 (10y)",false],
      ["WACC","레인지 (base 10.5%)","9.0 – 12.0%",false],
      ["Terminal growth (TGR)","","2.0%",false],
      ["Tax rate","US fed+state, AMPC 별도","23% (AMPC 비과세)",false],
      ["SKBA NOL (누적결손)","TCJA 80%한도·무기한","$1.5 bn",false],
      ["AMPC 적용기간","45X 법정 phase-out","2027–32 후 소멸",false],
      ["순부채 (DOE loan)","EV→Equity 브리지","$4.0 bn",false],
    ];
    dcfRows.forEach(r=>{ kv(rx,ry,r[0],r[1],r[2],r[3]); ry+=0.42; });
    s.addShape(pres.shapes.RECTANGLE, { x:rx, y:ry, w:colW2, h:0.42, fill:{color:"EFE9DC"}, line:{color:"EFE9DC",width:1} });
    s.addText("결론 EV (base)", { x:rx+0.1, y:ry, w:colW2*0.6, h:0.42, margin:0, fontFace:BODY, bold:true, fontSize:12, color:INK, valign:"middle" });
    s.addText("≈ $3.5 bn", { x:rx+colW2*0.45, y:ry, w:colW2*0.53, h:0.42, margin:0, align:"right", fontFace:MONO, bold:true, fontSize:12, color:INK, valign:"middle" });
    ry+=0.6;

    // sensitivity table
    s.addText("민감도 — EV ($bn) · WACC × TGR", { x:rx, y:ry, w:colW2, h:0.3, margin:0, fontFace:BODY, bold:true, fontSize:11.5, color:RED, charSpacing:1 });
    ry+=0.38;
    const sens=[
      [{t:"WACC ＼ TGR",h:true},{t:"1.5%",h:true},{t:"2.0%",h:true},{t:"2.5%",h:true}],
      [{t:"9.0%",rh:true},{t:"3.99"},{t:"4.11"},{t:"4.24"}],
      [{t:"10.5%",rh:true},{t:"3.38"},{t:"3.45",mid:true},{t:"3.53"}],
      [{t:"12.0%",rh:true},{t:"2.92"},{t:"2.97"},{t:"3.02"}],
    ];
    const stab = sens.map(row=>row.map(c=>({ text:c.t, options:{
      align:"center", fontFace:MONO, fontSize: c.h?10.5:12, bold: c.h||c.rh||c.mid,
      color: c.h?WHITE : c.mid?RED:INK,
      fill:{ color: c.h?INK : c.rh?"EFE9DC" : c.mid?"F6DDDF":WHITE },
      border:[{pt:1,color:LIND},{pt:1,color:LIND},{pt:1,color:LIND},{pt:1,color:LIND}], valign:"middle" } })));
    s.addTable(stab, { x:rx, y:ry, w:colW2, colW:[colW2*0.4,colW2*0.2,colW2*0.2,colW2*0.2], rowH:0.4 });
    ry+=0.4*4+0.18;

    const notes=[
      "DC블록 ASP $155·ex-AMPC EBITDA 9%·SKBA NOL $1.5bn(누적결손) 반영 → 초기 세금 절감으로 EV +$0.3bn.",
    ];
    notes.forEach(n=>{
      s.addText([{text:"– ",options:{color:RED,bold:true}},{text:n,options:{color:INK}}],
        { x:rx, y:ry, w:colW2, h:0.34, margin:0, fontFace:BODY, fontSize:10, valign:"top" });
      ry+=0.32;
    });

    // bottom full-width constraint note
    s.addText([
      { text:"제약이 현금흐름에 미치는 영향:  ", options:{ bold:true, color:RED } },
      { text:"공장2 DOE loan $4.0bn → 추가 federal funding 불가·부채상환이 FCF 선순위 점유. 20% CoC는 신규 equity를 자회사 prefs/mezz로 한정 → 하이브리드(HoldCo 우산+ring-fence) 구조로 반영. DCF EV $3.45bn(NOL 반영)은 ex-AMPC 멀티플 8~11x와 수렴.", options:{ color:MUTE } },
    ], { x:MX, y:Hh-1.05, w:W-2*MX, h:0.4, margin:0, fontFace:BODY, fontSize:10, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 4 — GO-TO-MARKET THEME (통합 안보 내러티브)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: INK };   // dark anchor slide
    s.addText([
      { text:"SK", options:{ fontFace:HEAD, bold:true, color:WHITE, fontSize:20, charSpacing:1 } },
      { text:"   GO-TO-MARKET · ANCHOR NARRATIVE", options:{ fontFace:BODY, bold:true, color:RED, fontSize:12, charSpacing:2 } },
    ], { x:MX, y:0.5, w:10.5, h:0.4, margin:0, valign:"middle" });
    s.addText("04 / 12", { x:W-2.3, y:0.5, w:1.6, h:0.4, margin:0, align:"right", fontFace:MONO, color:"A89F8E", fontSize:12, valign:"middle" });
    s.addText("시장에 내세울 단일 테마 — 「통합 안보(Integrated Security)」 내러티브",
      { x:MX, y:0.95, w:W-2*MX, h:0.6, margin:0, fontFace:HEAD, bold:true, color:WHITE, fontSize:25, valign:"middle" });
    s.addText("세 축을 하나로 묶어 '미국이 사야 하는 전략 자산'으로 포지셔닝 — 단순 배터리 공장이 아니라 안보 인프라",
      { x:MX, y:1.6, w:W-2*MX, h:0.45, margin:0, fontFace:BODY, color:"C9C1B2", fontSize:14, valign:"middle" });
    s.addShape(pres.shapes.LINE, { x:MX, y:2.2, w:W-2*MX, h:0, line:{ color:"4A443D", width:1 } });

    // three pillars
    const pillars=[
      [ic.shield,"FEOC-Clean 美 ESS 공급망","대중국 배제·FEOC 적격 셀을 미국 본토에서 생산. AI 데이터센터·그리드 ESS 수요를 비중국 공급으로 충족하는 유일 경로."],
      [ic.bolt,"에너지 안보 · 전력 인프라","brownfield·stranded power를 승계한 즉시 가동형 자산. 5년 인입 대기를 건너뛴 '전력'은 돈으로 살 수 없는 희소 자산."],
      [ic.micro,"AMPC 기반 美 현지화","IRA 45X 세액공제로 thin 마진을 보전. 미국 제조·고용·세수에 기여하는 '미국이 보조금을 주는' 구조."],
    ];
    const pw=(W-2*MX-2*0.4)/3;
    pillars.forEach((p,i)=>{
      const px=MX+i*(pw+0.4), py=2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:px, y:py, w:pw, h:2.35, rectRadius:0.06, fill:{color:"262019"}, line:{color:"3D372F",width:1} });
      s.addShape(pres.shapes.OVAL, { x:px+0.3, y:py+0.3, w:0.7, h:0.7, fill:{color:"332C24"}, line:{color:RED,width:1.5} });
      s.addImage({ data:p[0], x:px+0.47, y:py+0.47, w:0.36, h:0.36 });
      s.addText(p[1], { x:px+0.3, y:py+1.12, w:pw-0.6, h:0.55, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:14.5, valign:"top" });
      s.addText(p[2], { x:px+0.3, y:py+1.68, w:pw-0.6, h:0.6, margin:0, fontFace:BODY, color:"B5AD9E", fontSize:11, valign:"top" });
    });

    // convergence statement band
    const cy2=5.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:MX, y:cy2, w:W-2*MX, h:1.0, rectRadius:0.06, fill:{color:RED} });
    s.addText([
      { text:"세 축의 수렴 →  ", options:{ bold:true, color:WHITE, fontSize:15 } },
      { text:"\"중국 없는 AI·그리드 전력을 미국 땅에서 만드는 FEOC-clean 자산\"", options:{ color:WHITE, fontSize:15, italic:true } },
      { text:"  — 정부가 첫 번째 고객이자 투자자가 되는 논리", options:{ color:"FBE3E5", fontSize:13 } },
    ], { x:MX+0.4, y:cy2, w:W-2*MX-0.8, h:1.0, margin:0, fontFace:BODY, valign:"middle" });

    // why this beats single-theme
    const wy=6.55;
    s.addText("왜 단일 테마가 아니라 '묶음'인가", { x:MX, y:wy, w:W-2*MX, h:0.35, margin:0, fontFace:HEAD, bold:true, color:WHITE, fontSize:15 });
    const why=[
      ["FEOC만:","규제 변동·삼성/CATL 추격에 노출 → 단독으론 해자 약함"],
      ["AMPC만:","정치·세제 사이클 의존 → IRA 후퇴 시 가치 훼손"],
      ["전력만:","상품화된 부동산 논리 → 프리미엄 제한"],
      ["묶으면:","세 변수가 동시에 무너질 확률↓ → 정부 BATNA·우선권 확보의 명분이 됨"],
    ];
    const wh=(W-2*MX-0.4)/2;
    why.forEach((wd,i)=>{
      const wx=MX+(i%2)*(wh+0.4), wyy=wy+0.5+Math.floor(i/2)*0.62;
      s.addText([{text:wd[0]+" ",options:{bold:true,color: i===3?RED:"E4DCCD"}},{text:wd[1],options:{color:"B5AD9E"}}],
        { x:wx, y:wyy, w:wh, h:0.55, margin:0, fontFace:BODY, fontSize:11.5, valign:"middle" });
    });

    // footer (dark variant)
    s.addShape(pres.shapes.LINE, { x:MX, y:Hh-0.62, w:W-2*MX, h:0, line:{ color:"4A443D", width:0.75 } });
    s.addText([{ text:"Project [ ● ]", options:{ bold:true, color:"E4DCCD" } },{ text:"  ·  Strictly Private & Confidential", options:{ color:"A89F8E" } }],
      { x:MX, y:Hh-0.55, w:6, h:0.32, margin:0, fontFace:BODY, fontSize:10.5, valign:"middle" });
    s.addText("DRAFT — FOR DISCUSSION ONLY", { x:W/2-2.5, y:Hh-0.55, w:5, h:0.32, margin:0, align:"center", fontFace:BODY, bold:true, color:RED, fontSize:10.5, charSpacing:2, valign:"middle" });
    s.addText("2026. 6", { x:W-2.2, y:Hh-0.55, w:1.5, h:0.32, margin:0, align:"right", fontFace:MONO, color:"A89F8E", fontSize:10.5, valign:"middle" });
  }

  // ============================================================
  // SLIDE 5 — CRUCIBLE-STYLE FINANCING STRUCTURE
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING STRUCTURE · CRUCIBLE ANALOGUE", "05 / 12",
      [{ text:"Project Crucible형 파이낸싱 — 공장별 배분 & 투자자 모집", options:{} }],
      "통합 안보 내러티브 기반 · 미 연방정부 앵커 우선 · 공장1=전략지분 / 공장2=정부·부채 ring-fence");

    // ---- LEFT: Claude 추천 — 공장별 배분 구조 (분리 펀딩) ----
    const lx=MX, ly=2.7, lw=(W-2*MX-0.5)/2;
    s.addText([{text:"추천 구조  ",options:{bold:true,fontFace:HEAD,fontSize:16,color:INK}},
               {text:"— 공장별 '분리 펀딩' (split, not pooled)",options:{fontFace:BODY,fontSize:12,color:RED}}],
      { x:lx, y:ly, w:lw, h:0.35, margin:0, valign:"middle" });
    s.addShape(pres.shapes.LINE, { x:lx, y:ly+0.42, w:lw, h:0, line:{ color:INK, width:1.2 } });

    // HoldCo box
    let by=ly+0.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:lx+lw*0.2, y:by, w:lw*0.6, h:0.62, rectRadius:0.05, fill:{color:INK} });
    s.addText([{text:"SK 주도 US HoldCo / SPV\n",options:{bold:true,color:WHITE,fontSize:12.5,breakLine:true}},
               {text:"전략 + sovereign + financial co-invest",options:{color:"C9C1B2",fontSize:9}}],
      { x:lx+lw*0.2, y:by, w:lw*0.6, h:0.62, margin:0, align:"center", fontFace:BODY, valign:"middle" });
    // connectors
    by+=0.62;
    s.addShape(pres.shapes.LINE, { x:lx+lw*0.5, y:by, w:0, h:0.2, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:lx+lw*0.25, y:by+0.2, w:lw*0.5, h:0, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:lx+lw*0.25, y:by+0.2, w:0, h:0.18, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:lx+lw*0.75, y:by+0.2, w:0, h:0.18, line:{color:CMPD,width:1} });

    // two legs
    const legY=by+0.42, legW=lw*0.46;
    function leg(x,fill,title,tag,big,desc){
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:legY, w:legW, h:1.5, rectRadius:0.05, fill:{color:fill} });
      s.addText([{text:title,options:{bold:true,color:WHITE,fontSize:12}},{text:"   "+tag,options:{color:"EAD9DB",fontSize:8.5}}],
        { x:x+0.15, y:legY+0.12, w:legW-0.3, h:0.3, margin:0, fontFace:BODY, valign:"middle" });
      s.addText(big, { x:x+0.15, y:legY+0.42, w:legW-0.3, h:0.4, margin:0, fontFace:HEAD, bold:true, color:WHITE, fontSize:18 });
      s.addText(desc, { x:x+0.15, y:legY+0.85, w:legW-0.3, h:0.6, margin:0, fontFace:BODY, color:"F0E6E1", fontSize:9, valign:"top" });
    }
    leg(lx, P1, "공장 1 (모회사)","OPEN","≤ 20% equity","가동·AMPC-live. 외부지분 CoC 한도 내. 개조 $0.4bn은 project debt + AMPC 유동화.");
    leg(lx+lw*0.54, P2, "공장 2 (자회사)","RING-FENCED","Pref / Mezz","$4.0bn DOE loan·신규 연방자금 불가. $1.0bn은 자회사 비희석 prefs로, 모회사 cap table과 격리.");

    // allocation bar
    let ay=legY+1.55;
    s.addText("자금 배분 — capex는 공장2에 쏠리되, '지분 희석'은 쏠리면 안 됨", { x:lx, y:ay, w:lw, h:0.26, margin:0, fontFace:BODY, bold:true, fontSize:10.5, color:INK });
    ay+=0.33;
    s.addShape(pres.shapes.RECTANGLE, { x:lx, y:ay, w:lw*0.30, h:0.38, fill:{color:P1}, line:{color:P1,width:1} });
    s.addText("공장1 · ~$0.4bn", { x:lx, y:ay, w:lw*0.30, h:0.38, margin:0, align:"center", fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    s.addShape(pres.shapes.RECTANGLE, { x:lx+lw*0.30, y:ay, w:lw*0.70, h:0.38, fill:{color:P2}, line:{color:P2,width:1} });
    s.addText("공장2 · ~$1.0bn 개조", { x:lx+lw*0.30, y:ay, w:lw*0.70, h:0.38, margin:0, align:"center", fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    ay+=0.52;

    // two constraints callout
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:lx, y:ay, w:lw, h:1.42, rectRadius:0.05, fill:{color:WHITE}, line:{color:LIND,width:1}, shadow:sh() });
    s.addText("분리 펀딩을 강제하는 두 제약", { x:lx+0.18, y:ay+0.1, w:lw-0.36, h:0.26, margin:0, fontFace:BODY, bold:true, fontSize:11, color:RED, charSpacing:1 });
    const cons=[
      ["20% CoC 캡:","공장1 외부 equity 합산 ≤19.9% → DOE 코버넌트 보존. 무·제한의결권 클래스 활용."],
      ["연방자금 소진:","공장2는 신규 연방자금 stacking 불가 → $1.0bn은 자회사 prefs/mezz로 격리 조달."],
      ["AMPC = cross-collateral:","양 공장 AMPC 스트림을 담보·유동화해 보통주 희석 없이 커버리지 제고."],
    ];
    let cy3=ay+0.42;
    cons.forEach(c=>{
      s.addText([{text:"▸ ",options:{color:RED,bold:true}},{text:c[0]+" ",options:{bold:true,color:INK}},{text:c[1],options:{color:INK}}],
        { x:lx+0.18, y:cy3, w:lw-0.36, h:0.32, margin:0, fontFace:BODY, fontSize:9.2, valign:"top" });
      cy3+=0.31;
    });

    // ---- RIGHT: 투자자 모집 (정부 앵커 우선) ----
    const rx=MX+lw+0.5, rw=lw;
    s.addText([{text:"투자자 모집  ",options:{bold:true,fontFace:HEAD,fontSize:16,color:INK}},
               {text:"— 美 연방정부 앵커를 가장 먼저",options:{fontFace:BODY,fontSize:12,color:RED}}],
      { x:rx, y:2.7, w:rw, h:0.35, margin:0, valign:"middle" });
    s.addShape(pres.shapes.LINE, { x:rx, y:2.7+0.42, w:rw, h:0, line:{ color:INK, width:1.2 } });

    const wf=[
      [ic.flag,"1 · Sovereign 앵커 (최우선)","美 DoE/DoD/DFC 등 연방정부. Crucible처럼 정부 참여를 먼저 잠가 FAST-41·FEOC/CFIUS 옵틱을 정리하고 가격을 셋팅. 공장2 prefs를 타겟."],
      [ic.handsh,"2 · 전략적 / Offtake","하이퍼스케일러·ESS OEM·유틸리티가 공장1 ≤19.9% 소수지분 + 다년 LFP offtake. 매출 가시성으로 EV↑, CoC 미저촉."],
      [ic.coins,"3 · 인프라 / 크레딧","인프라-PE·사모 크레딧이 공장2 메자닌 흡수. AMPC + 계약 현금흐름으로 언더라이팅, 보통주 대비 구조적 선순위 → 모회사 무희석."],
      [ic.lock,"4 · 스폰서 (SK)","SK가 공장1 >80% + AMPC 경제권 보유. 신규 자금은 제한의결권·자회사 instrument에. DOE loan 상환 후 공장2 carve-out 옵션."],
    ];
    let wy2=2.7+0.6;
    const stepH=1.18;
    wf.forEach((w,i)=>{
      const yy=wy2+i*stepH;
      // connector line
      if(i<wf.length-1) s.addShape(pres.shapes.LINE, { x:rx+0.23, y:yy+0.5, w:0, h:stepH-0.05, line:{color:LIND,width:1.4} });
      s.addShape(pres.shapes.OVAL, { x:rx+0.08, y:yy+0.05, w:0.3, h:0.3, fill:{color:WHITE}, line:{color:RED,width:2} });
      s.addImage({ data:w[0], x:rx+0.135, y:yy+0.105, w:0.19, h:0.19 });
      s.addText(w[1], { x:rx+0.55, y:yy, w:rw-0.55, h:0.32, margin:0, fontFace:BODY, bold:true, fontSize:12.5, color:INK, valign:"middle" });
      s.addText(w[2], { x:rx+0.55, y:yy+0.32, w:rw-0.55, h:0.78, margin:0, fontFace:BODY, fontSize:9.7, color:"55504A", valign:"top" });
    });

    // bottom Crucible mapping note
    s.addText([
      { text:"Crucible 매핑:  ", options:{ bold:true, color:INK } },
      { text:"① 100% US HoldCo 신설 → ② 정부 주도 JV가 모회사 신주 제3자배정 인수(우호지분) → ③ 모회사 자체출자 합산해 HoldCo 출자 → ④ 정부 부채금융·보조금으로 총조달의 90%+를 정부·SI가 부담, SK는 운영주도권·기술 리더십 보유. 본 건은 DOE-encumbered 공장2를 ring-fence하여 연방 코버넌트와 모회사 equity 조달이 충돌하지 않게 분리.", options:{ color:MUTE } },
    ], { x:MX, y:Hh-1.25, w:W-2*MX, h:0.6, margin:0, fontFace:BODY, fontSize:10, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 6 — RING-FENCING 구조 친절 설명
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING DEEP-DIVE · RING-FENCING", "06 / 12",
      [{ text:"링펜싱(Ring-fencing) — 공장2를 '격벽'으로 격리하는 이유와 방법", options:{} }],
      "DOE $4.0bn 대출의 두 제약을 피하면서 자금을 넣는 구조 · 배의 격벽처럼 위험 전염 차단");

    // ---- 개념 정의 박스 (상단 가로) ----
    let ty=2.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:MX, y:ty, w:W-2*MX, h:0.92, rectRadius:0.05, fill:{color:"F2EEE7"}, line:{color:CMPD,width:1} });
    s.addText([{text:"링펜싱이란?  ",options:{bold:true,fontFace:HEAD,fontSize:14,color:RED}},
               {text:"한 그룹 안에서 특정 자산·부채를 ",options:{fontSize:12,color:INK}},
               {text:"별도 칸막이로 격리",options:{bold:true,fontSize:12,color:INK}},
               {text:"하는 구조. 배의 ",options:{fontSize:12,color:INK}},
               {text:"격벽(watertight bulkhead)",options:{bold:true,fontSize:12,color:INK}},
               {text:"과 같다 — 한 칸에 물이 새도 다른 칸은 멀쩡해 배 전체는 가라앉지 않는다.",options:{fontSize:12,color:INK}}],
      { x:MX+0.25, y:ty+0.12, w:W-2*MX-0.5, h:0.32, margin:0, fontFace:BODY, valign:"middle" });
    s.addText([{text:"이 딜에서는 → ",options:{bold:true,fontSize:11,color:RED}},
               {text:"공장2(SKOT)에 묶인 $4.0bn DOE 대출의 까다로운 조건이 모회사·공장1로 번지지 않도록 공장2를 별도 칸에 가둔다.",options:{fontSize:11,color:"55504A"}}],
      { x:MX+0.25, y:ty+0.5, w:W-2*MX-0.5, h:0.32, margin:0, fontFace:BODY, valign:"middle" });

    // ---- LEFT: 왜 필요한가 — 두 제약 ----
    const lx=MX, lyy=3.78, lw=(W-2*MX-0.5)/2;
    s.addText("왜 필요한가 — DOE 대출의 두 제약", { x:lx, y:lyy, w:lw, h:0.32, margin:0, fontFace:HEAD, bold:true, fontSize:14, color:INK });
    s.addShape(pres.shapes.LINE, { x:lx, y:lyy+0.4, w:lw, h:0, line:{color:INK,width:1.2} });

    const probs=[
      ["①","연방자금 중복(stacking) 제한","공장2는 이미 DOE 정부자금을 받음 → 추가 연방자금을 더 못 끌어옴. '정부 돈은 이미 가득 찬' 상태."],
      ["②","Change-of-Control 코버넌트","소유주가 일정% 이상 바뀌면 대출 즉시 회수·재협상. → 외부에 지분을 함부로 못 팖 (공장1 20% 초과 트리거)."],
    ];
    let py=lyy+0.56;
    probs.forEach(p=>{
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:lx, y:py, w:lw, h:1.08, rectRadius:0.05, fill:{color:WHITE}, line:{color:LIND,width:1}, shadow:sh() });
      s.addShape(pres.shapes.OVAL, { x:lx+0.16, y:py+0.2, w:0.52, h:0.52, fill:{color:RED} });
      s.addText(p[0], { x:lx+0.16, y:py+0.2, w:0.52, h:0.52, margin:0, align:"center", valign:"middle", fontFace:HEAD, bold:true, color:WHITE, fontSize:20 });
      s.addText(p[1], { x:lx+0.85, y:py+0.14, w:lw-1.0, h:0.32, margin:0, fontFace:BODY, bold:true, fontSize:12, color:INK, valign:"middle" });
      s.addText(p[2], { x:lx+0.85, y:py+0.46, w:lw-1.0, h:0.56, margin:0, fontFace:BODY, fontSize:9.8, color:"55504A", valign:"top" });
      py+=1.2;
    });
    // 문제 요약
    s.addText([{text:"딜레마:  ",options:{bold:true,color:RED}},
               {text:"외부 투자금이 필요한데 — 공장2에 직접 받으면 코버넌트 위반, 모회사 전체로 받으면 공장2 부채가 전염.",options:{color:INK}}],
      { x:lx, y:py+0.04, w:lw, h:0.5, margin:0, fontFace:BODY, fontSize:10.5, valign:"top" });

    // ---- RIGHT: 어떻게 격리하나 — 도구 ----
    const rx=MX+lw+0.5, rw=lw;
    s.addText("어떻게 격리하나 — '지배구조 안 바꾸는' 자금", { x:rx, y:lyy, w:rw, h:0.32, margin:0, fontFace:HEAD, bold:true, fontSize:14, color:INK });
    s.addShape(pres.shapes.LINE, { x:rx, y:lyy+0.4, w:rw, h:0, line:{color:INK,width:1.2} });

    // 자금 도구 비교 표
    const tools=[
      ["보통주 (equity)","의결권+소유권","CoC 트리거·모회사 희석","✗",RED],
      ["우선주 (preferred)","배당우선·의결권 제한","소유권 변동 없음→CoC 회피","✓","1E7A4D"],
      ["메자닌 (mezz)","부채+지분 중간","자회사 레벨 격리·모회사 무영향","✓","1E7A4D"],
    ];
    let tyy=lyy+0.56;
    // header row
    s.addShape(pres.shapes.RECTANGLE, { x:rx, y:tyy, w:rw, h:0.34, fill:{color:INK} });
    s.addText("자금 도구", { x:rx+0.1, y:tyy, w:rw*0.28, h:0.34, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    s.addText("성격", { x:rx+rw*0.28, y:tyy, w:rw*0.26, h:0.34, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    s.addText("링펜싱 효과", { x:rx+rw*0.54, y:tyy, w:rw*0.38, h:0.34, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    s.addText("가능", { x:rx+rw*0.92, y:tyy, w:rw*0.08, h:0.34, margin:0, align:"center", fontFace:BODY, bold:true, color:WHITE, fontSize:9.5, valign:"middle" });
    tyy+=0.34;
    tools.forEach((t,i)=>{
      const rh=0.5;
      s.addShape(pres.shapes.RECTANGLE, { x:rx, y:tyy, w:rw, h:rh, fill:{color: i%2?"F2EEE7":"FBF9F5"}, line:{color:LIND,width:0.5} });
      s.addText(t[0], { x:rx+0.1, y:tyy, w:rw*0.28, h:rh, margin:0, fontFace:BODY, bold:true, fontSize:9.8, color:INK, valign:"middle" });
      s.addText(t[1], { x:rx+rw*0.28, y:tyy, w:rw*0.26, h:rh, margin:0, fontFace:BODY, fontSize:9, color:"55504A", valign:"middle" });
      s.addText(t[2], { x:rx+rw*0.54, y:tyy, w:rw*0.38, h:rh, margin:0, fontFace:BODY, fontSize:9, color:"55504A", valign:"middle" });
      s.addText(t[3], { x:rx+rw*0.92, y:tyy, w:rw*0.08, h:rh, margin:0, align:"center", fontFace:HEAD, bold:true, fontSize:15, color:t[4], valign:"middle" });
      tyy+=rh;
    });
    // 핵심 메시지
    s.addText([{text:"→ 핵심:  ",options:{bold:true,color:RED}},
               {text:"우선주·메자닌은 '돈은 받되 지배구조는 안 바뀌는' 도구. DOE 코버넌트를 안 건드리며, 자금이 공장2 칸 안에만 머물러 모회사·공장1로 전염되지 않는다.",options:{color:INK}}],
      { x:rx, y:tyy+0.1, w:rw, h:0.7, margin:0, fontFace:BODY, fontSize:10, valign:"top" });
    // AMPC cross-collateral 보조
    s.addText([{text:"＋ AMPC 교차담보:  ",options:{bold:true,color:RED}},
               {text:"양 공장 AMPC 현금흐름을 한데 묶어 담보·유동화 → 보통주 희석 없이 커버리지 제고.",options:{color:"55504A"}}],
      { x:rx, y:tyy+0.85, w:rw, h:0.5, margin:0, fontFace:BODY, fontSize:9.8, valign:"top" });

    footer(s);
  }

  // ============================================================
  // SLIDE 7 — CRUCIBLE 비교 (차용 vs 차별)
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING · CRUCIBLE BENCHMARK", "07 / 12",
      [{ text:"고려아연 Crucible 대비 — 무엇을 차용하고 무엇을 다르게 했나", options:{} }],
      "Crucible: 美 정부 JV가 고려아연 본체 지분 ~10% 제3자배정 · 정부·FI 90%+ 부담 · CHIPS grant 프로젝트 직접출자");

    // 차용 3개 (좌) / 차별 2개 (우)
    const colW=(W-2*MX-0.5)/2, lx=MX, rx=MX+colW+0.5, ty=2.7;
    // 좌: 차용
    s.addShape(pres.shapes.RECTANGLE, { x:lx, y:ty, w:colW, h:0.44, fill:{color:INK} });
    s.addText("✓ Crucible에서 차용한 골격", { x:lx+0.15, y:ty, w:colW-0.3, h:0.44, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:13, valign:"middle" });
    const borrow=[
      ["정부 앵커 우선","DoE/DoD/DFC를 먼저 잠가 FAST-41·FEOC·CFIUS 옵틱 정리·가격 셋팅"],
      ["정부·SI 90%+ 부담","SK 잔여출자 → 부채비율 부담 없이 재무안정성 개선, 운영주도권 유지"],
      ["HoldCo/SPV 다층","자금 성격별로 칸을 나눠 받는 골격 — KZ→HoldCo→프로젝트 구조"],
    ];
    let by=ty+0.6;
    borrow.forEach(b=>{
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:lx, y:by, w:colW, h:1.05, rectRadius:0.04, fill:{color:WHITE}, line:{color:LIND,width:1}, shadow:sh() });
      s.addText(b[0], { x:lx+0.18, y:by+0.13, w:colW-0.36, h:0.32, margin:0, fontFace:BODY, bold:true, fontSize:12.5, color:INK, valign:"middle" });
      s.addText(b[1], { x:lx+0.18, y:by+0.46, w:colW-0.36, h:0.52, margin:0, fontFace:BODY, fontSize:10, color:"55504A", valign:"top" });
      by+=1.17;
    });
    // 우: 차별
    s.addShape(pres.shapes.RECTANGLE, { x:rx, y:ty, w:colW, h:0.44, fill:{color:RED} });
    s.addText("✗ Crucible과 결정적으로 다른 점", { x:rx+0.15, y:ty, w:colW-0.3, h:0.44, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:13, valign:"middle" });
    const diff=[
      ["제3자배정 위치","Crucible: 고려아연 본체(모회사) 지분 10% → 영풍·MBK 거버넌스 반발","우리: 모회사 아닌 공장1 자회사 ≤19.9%만 → 모회사 의결권 희석 회피"],
      ["링펜싱 추가","Crucible: 신설 프로젝트라 격리할 기존 부채 없음","우리: 공장2에 DOE $4bn 기존 묶임 → ring-fence로 전염 차단 (고유 장치)"],
      ["보조금 성격","Crucible: CHIPS = 일회성 설비보조금, 프로젝트 직접출자","우리: AMPC = 생산세액공제, 매년 현금흐름 → cross-collateral 유동화"],
    ];
    let dy=ty+0.6;
    diff.forEach(d=>{
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:rx, y:dy, w:colW, h:1.05, rectRadius:0.04, fill:{color:"FBF3F2"}, line:{color:RED,width:1}, shadow:sh() });
      s.addText(d[0], { x:rx+0.18, y:dy+0.1, w:colW-0.36, h:0.28, margin:0, fontFace:BODY, bold:true, fontSize:12, color:RED, valign:"middle" });
      s.addText([{text:"KZ: ",options:{bold:true,color:MUTE}},{text:d[1].replace(/^Crucible: /,''),options:{color:"55504A"}}], { x:rx+0.18, y:dy+0.4, w:colW-0.36, h:0.3, margin:0, fontFace:BODY, fontSize:9, valign:"top" });
      s.addText([{text:"SK: ",options:{bold:true,color:RED}},{text:d[2].replace(/^우리: /,''),options:{color:INK}}], { x:rx+0.18, y:dy+0.68, w:colW-0.36, h:0.32, margin:0, fontFace:BODY, fontSize:9, valign:"top" });
      dy+=1.17;
    });
    s.addText([{text:"요약:  ",options:{bold:true,color:RED}},
      {text:"Crucible의 '정부앵커+HoldCo+정부90%+운영주도권' 골격을 차용하되, ① 외국 의결권을 모회사가 아닌 자회사(공장1 ≤19.9%)에 한정해 거버넌스 논란 회피, ② Crucible엔 없던 링펜싱(공장2 DOE 격리)을 추가한 변형.",options:{color:INK}}],
      { x:MX, y:Hh-1.05, w:W-2*MX, h:0.6, margin:0, fontFace:BODY, fontSize:10.5, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 8 — 링펜싱 계약 스택 (5층)
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING · RING-FENCING CONTRACT STACK", "08 / 12",
      [{ text:"링펜싱을 계약서에 구현하는 법 — 5개 층이 맞물려 격벽 완성", options:{} }],
      "단일 조항 아님 · 법인구조 → LLC약정 → 자금계약 → 채권자간약정 → 담보계약 · 전체 관통 = non-recourse");

    const layers=[
      ["①","법인 분리 (Entity Separation)","별도 SPV(SKOT HoldCo→OpCo) · bankruptcy-remote · 독립이사(Independent Director) 동의 없인 도산신청 불가","격벽의 뼈대"],
      ["②","LLC 약정 (Operating Agreement)","Separateness 서약(자산혼화 금지·별도장부) + Restricted Payments(커버리지 충족 후만 모회사 배당) + Permitted Debt 한정","칸막이 규칙"],
      ["③","자금조달 계약","우선주 양수계약: 의결권 제한(CoC 회피)·10% 우선배당·청산우선권 / 메자닌 대출계약 / DOE Loan(이자-only)","지배구조 안 건드림"],
      ["④","채권자 간 약정 (Intercreditor)","Subordination(메자닌·SK 후순위) + Payment Waterfall(운영비→DOE이자→우선배당→메자닌→SK) + Standstill","워터폴의 법적 실체"],
      ["⑤","담보계약 (Security Agreement)","45X AMPC 수취권리(receivables)를 담보 제공·유동화(IRA 양도) → 미래현금 당겨오고 대주에 담보","AMPC 교차담보"],
    ];
    let ly=2.65; const lh=1.04, lw=W-2*MX;
    layers.forEach((L,i)=>{
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:MX, y:ly, w:lw, h:lh-0.12, rectRadius:0.04, fill:{color: i%2?"F2EEE7":WHITE}, line:{color:LIND,width:1}, shadow:sh() });
      // 번호 원
      s.addShape(pres.shapes.OVAL, { x:MX+0.18, y:ly+0.22, w:0.5, h:0.5, fill:{color:RED} });
      s.addText(L[0], { x:MX+0.18, y:ly+0.22, w:0.5, h:0.5, margin:0, align:"center", valign:"middle", fontFace:HEAD, bold:true, color:WHITE, fontSize:17 });
      s.addText(L[1], { x:MX+0.85, y:ly+0.13, w:3.3, h:0.66, margin:0, fontFace:BODY, bold:true, fontSize:13, color:INK, valign:"middle" });
      s.addText(L[2], { x:MX+4.2, y:ly+0.13, w:lw-5.7, h:0.66, margin:0, fontFace:BODY, fontSize:9.8, color:"55504A", valign:"middle" });
      s.addText(L[3], { x:MX+lw-1.5, y:ly+0.13, w:1.4, h:0.66, margin:0, align:"right", fontFace:BODY, bold:true, italic:true, fontSize:9.5, color:RED, valign:"middle" });
      ly+=lh;
    });
    // non-recourse 강조 바
    s.addShape(pres.shapes.RECTANGLE, { x:MX, y:ly+0.04, w:lw, h:0.5, fill:{color:INK} });
    s.addText([{text:"전체 관통 핵심 = Non-recourse:  ",options:{bold:true,color:WHITE,fontSize:12}},
      {text:"\"이 채무는 SKOT SPV 자산에만 청구, 모회사(SK)·공장1(SKBA)엔 소구하지 않는다\" — 이 한 줄이 격벽의 본질.",options:{color:"EAD9DB",fontSize:11}}],
      { x:MX+0.2, y:ly+0.04, w:lw-0.4, h:0.5, margin:0, fontFace:BODY, valign:"middle" });
    footer(s);
  }

  // ============================================================
  // SLIDE 9 — 하이브리드 최적구조 + waiver (3옵션)
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING · OPTIMAL STRUCTURE", "09 / 12",
      [{ text:"최적 구조 — 'HoldCo 우산 + 법인별 격벽' 하이브리드", options:{} }],
      "DOE 사전 waiver 전제 · waiver는 통합의 열쇠 아닌 '분리 구조의 좁은 미세조정'으로만 사용");

    // 3 옵션 비교 (좌)
    const colW=(W-2*MX-0.5)*0.52, lx=MX, ty=2.7;
    s.addText("세 가지 구조 비교", { x:lx, y:ty, w:colW, h:0.32, margin:0, fontFace:HEAD, bold:true, fontSize:14, color:INK });
    s.addShape(pres.shapes.LINE, { x:lx, y:ty+0.4, w:colW, h:0, line:{color:INK,width:1.2} });
    const opts=[
      ["(A) 완전 분리","현재 구조 · 각자 조달·OT 격리","전염 차단 확실·waiver 불요","OT 조달비 비쌈·자본효율 낮음",CMP],
      ["(B) 하이브리드 ★","HoldCo 통합조달 + SPV 격벽","비용↓ + 격벽 유지 (최적)","waiver 협상 필요",RED],
      ["(C) 완전 통합","Crucible형 단일 HoldCo","자본효율 최고","담보오염·전염부활·exit경직",CMP],
    ];
    let oy=ty+0.56;
    opts.forEach(o=>{
      const isBest=o[4]===RED;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:lx, y:oy, w:colW, h:1.15, rectRadius:0.04, fill:{color:isBest?"FBF3F2":WHITE}, line:{color:isBest?RED:LIND,width:isBest?1.5:1}, shadow:sh() });
      s.addText(o[0], { x:lx+0.16, y:oy+0.1, w:colW-0.3, h:0.3, margin:0, fontFace:BODY, bold:true, fontSize:12.5, color:isBest?RED:INK, valign:"middle" });
      s.addText(o[1], { x:lx+0.16, y:oy+0.4, w:colW-0.3, h:0.26, margin:0, fontFace:BODY, fontSize:9.5, color:"55504A", valign:"top" });
      s.addText([{text:"＋ ",options:{color:POS,bold:true}},{text:o[2],options:{color:INK}}], { x:lx+0.16, y:oy+0.66, w:colW-0.3, h:0.24, margin:0, fontFace:BODY, fontSize:9, valign:"top" });
      s.addText([{text:"− ",options:{color:RED,bold:true}},{text:o[3],options:{color:"55504A"}}], { x:lx+0.16, y:oy+0.88, w:colW-0.3, h:0.24, margin:0, fontFace:BODY, fontSize:9, valign:"top" });
      oy+=1.27;
    });

    // 하이브리드 구조도 (우)
    const rx=MX+colW+0.5, rw=W-rx-MX;
    s.addText("왜 waiver 받아도 '분리' 유지가 유리한가", { x:rx, y:ty, w:rw, h:0.32, margin:0, fontFace:HEAD, bold:true, fontSize:14, color:INK });
    s.addShape(pres.shapes.LINE, { x:rx, y:ty+0.4, w:rw, h:0, line:{color:INK,width:1.2} });
    // HoldCo box
    let dgy=ty+0.58;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:rx+rw*0.2, y:dgy, w:rw*0.6, h:0.56, rectRadius:0.04, fill:{color:INK} });
    s.addText([{text:"SK 주도 US HoldCo\n",options:{bold:true,color:WHITE,fontSize:11,breakLine:true}},{text:"공통 자금 통합 조달(저리)",options:{color:"C9C1B2",fontSize:8.5}}], { x:rx+rw*0.2, y:dgy, w:rw*0.6, h:0.56, margin:0, align:"center", fontFace:BODY, valign:"middle" });
    dgy+=0.56;
    s.addShape(pres.shapes.LINE, { x:rx+rw*0.5, y:dgy, w:0, h:0.16, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:rx+rw*0.25, y:dgy+0.16, w:rw*0.5, h:0, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:rx+rw*0.25, y:dgy+0.16, w:0, h:0.14, line:{color:CMPD,width:1} });
    s.addShape(pres.shapes.LINE, { x:rx+rw*0.75, y:dgy+0.16, w:0, h:0.14, line:{color:CMPD,width:1} });
    dgy+=0.32;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:rx, y:dgy, w:rw*0.46, h:0.66, rectRadius:0.04, fill:{color:P1} });
    s.addText([{text:"BA SPV\n",options:{bold:true,color:WHITE,fontSize:10.5,breakLine:true}},{text:"DOE-clean·외부지분≤19.9%",options:{color:"EAD9DB",fontSize:8}}], { x:rx, y:dgy, w:rw*0.46, h:0.66, margin:0, align:"center", fontFace:BODY, valign:"middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:rx+rw*0.54, y:dgy, w:rw*0.46, h:0.66, rectRadius:0.04, fill:{color:P2} });
    s.addText([{text:"OT SPV\n",options:{bold:true,color:WHITE,fontSize:10.5,breakLine:true}},{text:"DOE ring-fenced·우선주/메자닌",options:{color:"EAD9DB",fontSize:8}}], { x:rx+rw*0.54, y:dgy, w:rw*0.46, h:0.66, margin:0, align:"center", fontFace:BODY, valign:"middle" });
    dgy+=0.82;
    // 3 이유
    const reasons=[
      ["담보 오염 방지","DOE 담보가 OT에만 → BA는 DOE-clean 유지 → BA 조달 더 저렴"],
      ["FEOC/AMPC 적격성","적격성은 법인별 판정 → 분리하면 각자 깨끗하게 판정"],
      ["Exit 유연성","BA만 매각 / OT만 DOE상환 후 carve-out 가능 (옵션가치 큼)"],
    ];
    reasons.forEach((rr,i)=>{
      s.addText([{text:"▸ "+rr[0]+":  ",options:{bold:true,color:RED}},{text:rr[1],options:{color:INK}}],
        { x:rx, y:dgy+i*0.42, w:rw, h:0.4, margin:0, fontFace:BODY, fontSize:9.5, valign:"top" });
    });

    s.addText([{text:"waiver의 최적 용도:  ",options:{bold:true,color:RED}},
      {text:"완전 통합으로 가는 열쇠가 아니라 — HoldCo 우산에서 공통 자금을 모으되 OT로는 DOE가 허락한 안전한 형태(비담보 후순위)로만 흘려보내는, 분리 구조의 미세조정.",options:{color:INK}}],
      { x:MX, y:Hh-0.92, w:W-2*MX, h:0.5, margin:0, fontFace:BODY, fontSize:10.5, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 10 — 공장별 Financial Projection
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCIAL PROJECTION · BY ENTITY", "10 / 12",
      [{ text:"공장별 재무 프로젝션 — SKBA · SKOT · 보통주(SK) FCFE", options:{} }],
      "DC블록 $155·AMPC $45·ex-AMPC EBITDA 9%·SKBA NOL $1.5bn 반영 · 하이브리드 자본구조 · $mm");

    const yrs=["2027","2028","2029","2030","2031","2032","2033"];
    // 데이터 (모델 추출)
    const D={
      ba_rev:[723,1643,1946,2177,2249,2244,2213], ba_ebit:[47,122,148,169,175,174,171], ba_ampc:[210,477,572,487,340,172,0], ba_fcff:[-30,394,638,576,456,299,130], ba_fcfe:[-76,348,593,530,410,254,84],
      ot_rev:[0,1242,1927,2420,2729,2861,2822], ot_ebit:[-47,45,107,151,179,191,187], ot_ampc:[0,374,588,561,428,227,0], ot_fcff:[-500,-40,582,598,504,335,132], ot_fcfe:[-723,-262,109,142,66,-85,-21],
      sk_fcfe:[-798,86,702,673,477,169,63],
    };
    function miniTable(x,y,w,title,tint,rows){
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h:0.4, fill:{color:tint} });
      s.addText(title, { x:x+0.12, y, w:w-0.24, h:0.4, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:12.5, valign:"middle" });
      // year header
      let yy=y+0.4;
      const lblW=1.85, cw=(w-lblW)/7;
      s.addShape(pres.shapes.RECTANGLE, { x, y:yy, w, h:0.32, fill:{color:"F2EEE7"} });
      s.addText("($mm)", { x:x+0.1, y:yy, w:lblW-0.1, h:0.32, margin:0, fontFace:BODY, fontSize:8.5, color:MUTE, valign:"middle" });
      yrs.forEach((yr,i)=>s.addText(yr, { x:x+lblW+i*cw, y:yy, w:cw, h:0.32, margin:0, align:"center", fontFace:MONO, bold:true, fontSize:9, color:INK, valign:"middle" }));
      yy+=0.32;
      rows.forEach((rw,ri)=>{
        const vals=D[rw[1]], isBold=rw[2];
        s.addShape(pres.shapes.RECTANGLE, { x, y:yy, w, h:0.34, fill:{color: ri%2?"FBF9F5":"FFFFFF"}, line:{color:LINE,width:0.5} });
        s.addText(rw[0], { x:x+0.1, y:yy, w:lblW-0.1, h:0.34, margin:0, fontFace:BODY, bold:isBold, fontSize:9, color:INK, valign:"middle" });
        vals.forEach((v,i)=>{
          const neg=v<0;
          s.addText(neg?"("+Math.abs(v)+")":String(v), { x:x+lblW+i*cw, y:yy, w:cw, h:0.34, margin:0, align:"center", fontFace:MONO, bold:isBold, fontSize:8.5, color: neg?RED:(isBold?INK:"55504A"), valign:"middle" });
        });
        yy+=0.34;
      });
      return yy;
    }
    const tW=W-2*MX;
    let cy2=2.62;
    // SKBA
    cy2=miniTable(MX, cy2, tW, "공장 1 — SKBA (조지아) · LFP 파우치 21.2GWh", P1, [
      ["매출","ba_rev",false],["EBIT (pre-AMPC)","ba_ebit",false],["AMPC 45X","ba_ampc",false],["Unlevered FCFF","ba_fcff",true],["SKBA FCFE","ba_fcfe",true],
    ]);
    cy2+=0.22;
    // SKOT
    cy2=miniTable(MX, cy2, tW, "공장 2 — SKOT (테네시) · 각형+파우치 29.7GWh", P2, [
      ["매출","ot_rev",false],["EBIT (pre-AMPC)","ot_ebit",false],["AMPC 45X","ot_ampc",false],["Unlevered FCFF","ot_fcff",true],["SKOT FCFE (보통주)","ot_fcfe",true],
    ]);
    cy2+=0.22;
    // SK 합산 FCFE (1줄 강조)
    s.addShape(pres.shapes.RECTANGLE, { x:MX, y:cy2, w:tW, h:0.46, fill:{color:INK} });
    const lblW=1.85, cw=(tW-lblW)/7;
    s.addText("보통주(SK) FCFE 합계", { x:MX+0.1, y:cy2, w:lblW-0.1, h:0.46, margin:0, fontFace:BODY, bold:true, color:WHITE, fontSize:10, valign:"middle" });
    D.sk_fcfe.forEach((v,i)=>{
      const neg=v<0;
      s.addText(neg?"("+Math.abs(v)+")":String(v), { x:MX+lblW+i*cw, y:cy2, w:cw, h:0.46, margin:0, align:"center", fontFace:MONO, bold:true, fontSize:10, color: neg?"FF9B9B":"FFFFFF", valign:"middle" });
    });

    s.addText([{text:"읽는 법:  ",options:{bold:true,color:INK}},
      {text:"ex-AMPC EBITDA 9%로 양사 FCFF 견조. 메자닌 상환(2029~32) 부담으로 SKOT FCFE는 후기 소폭 적자, 그러나 DSCR 1.4~1.6x로 메자닌 적격. SKBA(DOE-clean)는 전구간 흑자. 보통주(SK) FCFE 합계는 2028년 흑자전환.",options:{color:"55504A"}}],
      { x:MX, y:Hh-0.82, w:W-2*MX, h:0.5, margin:0, fontFace:BODY, fontSize:9.5, valign:"top" });
    footer(s);
  }

  // ============================================================
  // SLIDE 11 — 링펜싱 구조도 + 지키는 조항
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING · RING-FENCING STRUCTURE MAP", "11 / 12",
      [{ text:"링펜싱 구조도 — 무엇을 어떻게 격리하고, 무슨 조항으로 지키나", options:{} }],
      "HoldCo 우산 + 법인별 격벽 · 자금 흐름과 격리 경계 · 각 경계를 지키는 계약 조항");

    // ===== 좌: 구조도 =====
    const lx=MX, lw=(W-2*MX)*0.56, ty=2.65;
    // HoldCo
    const hcW=2.8, hcX=lx+(lw-hcW)/2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:hcX, y:ty, w:hcW, h:0.72, rectRadius:0.05, fill:{color:INK} });
    s.addText([{text:"SK 주도 US HoldCo\n",options:{bold:true,color:WHITE,fontSize:12,breakLine:true}},{text:"정부앵커·전략·인프라크레딧 통합조달",options:{color:"C9C1B2",fontSize:8}}],
      { x:hcX, y:ty, w:hcW, h:0.72, margin:0, align:"center", fontFace:BODY, valign:"middle" });
    // 분기선
    const midY=ty+0.72;
    s.addShape(pres.shapes.LINE, { x:hcX+hcW/2, y:midY, w:0, h:0.22, line:{color:CMPD,width:1.5} });
    const baX=lx+0.1, otX=lx+lw-2.7-0.1, boxW=2.7, boxY=midY+0.5, boxH=1.5;
    s.addShape(pres.shapes.LINE, { x:baX+boxW/2, y:midY+0.22, w:(otX+boxW/2)-(baX+boxW/2), h:0, line:{color:CMPD,width:1.5} });
    s.addShape(pres.shapes.LINE, { x:baX+boxW/2, y:midY+0.22, w:0, h:0.28, line:{color:CMPD,width:1.5} });
    s.addShape(pres.shapes.LINE, { x:otX+boxW/2, y:midY+0.22, w:0, h:0.28, line:{color:CMPD,width:1.5} });
    s.addText("후순위 출자 $300", { x:baX, y:midY+0.18, w:boxW, h:0.26, margin:0, align:"center", fontFace:BODY, fontSize:7.5, color:MUTE, italic:true });
    s.addText("후순위 출자 $500 (waiver)", { x:otX, y:midY+0.18, w:boxW, h:0.26, margin:0, align:"center", fontFace:BODY, fontSize:7.5, color:MUTE, italic:true });
    // BA SPV
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:baX, y:boxY, w:boxW, h:boxH, rectRadius:0.05, fill:{color:WHITE}, line:{color:P1,width:2} });
    s.addShape(pres.shapes.RECTANGLE, { x:baX, y:boxY, w:boxW, h:0.42, fill:{color:P1} });
    s.addText("BA SPV (SKBA)", { x:baX, y:boxY, w:boxW, h:0.42, margin:0, align:"center", fontFace:BODY, bold:true, color:WHITE, fontSize:11, valign:"middle" });
    s.addText([
      {text:"DOE-clean ✓\n",options:{color:POS,bold:true,fontSize:9,breakLine:true}},
      {text:"외부지분 ≤19.9%\n",options:{color:INK,fontSize:9,breakLine:true}},
      {text:"project debt $400\n",options:{color:"55504A",fontSize:9,breakLine:true}},
      {text:"파우치 21.2GWh",options:{color:"55504A",fontSize:9}},
    ], { x:baX+0.15, y:boxY+0.5, w:boxW-0.3, h:boxH-0.6, margin:0, fontFace:BODY, valign:"top", lineSpacingMultiple:1.1 });
    // OT SPV
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:otX, y:boxY, w:boxW, h:boxH, rectRadius:0.05, fill:{color:WHITE}, line:{color:P2,width:2} });
    s.addShape(pres.shapes.RECTANGLE, { x:otX, y:boxY, w:boxW, h:0.42, fill:{color:P2} });
    s.addText("OT SPV (SKOT)", { x:otX, y:boxY, w:boxW, h:0.42, margin:0, align:"center", fontFace:BODY, bold:true, color:WHITE, fontSize:11, valign:"middle" });
    s.addText([
      {text:"DOE ring-fenced 🔒\n",options:{color:RED,bold:true,fontSize:9,breakLine:true}},
      {text:"DOE loan $4,000 (이자만)\n",options:{color:INK,fontSize:9,breakLine:true}},
      {text:"메자닌 $1,000 (7%)\n",options:{color:"55504A",fontSize:9,breakLine:true}},
      {text:"각형+파우치 29.7GWh",options:{color:"55504A",fontSize:9}},
    ], { x:otX+0.15, y:boxY+0.5, w:boxW-0.3, h:boxH-0.6, margin:0, fontFace:BODY, valign:"top", lineSpacingMultiple:1.1 });
    // non-recourse 경계 표시
    const nrY=boxY+boxH+0.25;
    s.addShape(pres.shapes.RECTANGLE, { x:lx, y:nrY, w:lw, h:0.5, fill:{color:"F2EEE7"}, line:{color:RED,width:1.5,dashType:"dash"} });
    s.addText([{text:"⊗ Non-recourse 격벽:  ",options:{bold:true,color:RED,fontSize:10}},{text:"각 SPV 채무는 자기 자산에만 청구 · 모회사·타 SPV로 전염 차단",options:{color:INK,fontSize:9.5}}],
      { x:lx+0.15, y:nrY, w:lw-0.3, h:0.5, margin:0, fontFace:BODY, valign:"middle" });
    s.addText("자금: 위(HoldCo)에서 통합조달 → 아래(SPV)로 후순위 배분 · 자산: 각 SPV에 격리",
      { x:lx, y:nrY+0.6, w:lw, h:0.3, margin:0, fontFace:BODY, fontSize:8.5, color:MUTE, italic:true, align:"center" });

    // ===== 우: 경계별 지키는 조항 =====
    const rx=lx+lw+0.4, rw=W-rx-MX;
    s.addText("격리 경계를 지키는 계약 조항", { x:rx, y:ty, w:rw, h:0.32, margin:0, fontFace:HEAD, bold:true, fontSize:13.5, color:INK });
    s.addShape(pres.shapes.LINE, { x:rx, y:ty+0.4, w:rw, h:0, line:{color:INK,width:1.2} });
    const clauses=[
      ["법인 분리","Bankruptcy-remote SPV + 독립이사 — 모회사가 일방 도산·자산이전 못 함"],
      ["Separateness","자산혼화 금지·별도 장부/계좌 — veil-piercing 방어"],
      ["Restricted Payments","DOE이자·메자닌 충족 + DSCR 1.2x 후에만 모회사 배당"],
      ["Permitted Debt","DOE+승인 메자닌 외 추가 부채 금지 — 칸 안 부채 총량 통제"],
      ["의결권 제한","메자닌·우선주 무의결권 → OT 지배구조 불변 → DOE CoC 회피"],
      ["Intercreditor","워터폴(운영비→DOE→메자닌→SK)+Standstill — 후순위 회수순서 고정"],
      ["AMPC Security","45X 수취권리 담보·유동화 — 메자닌 상환재원 확보"],
      ["Non-recourse","모회사·타 SPV 무보증 — 격벽의 최종 방어선"],
    ];
    let cyy=ty+0.54; const chh=0.62;
    clauses.forEach((c,i)=>{
      s.addShape(pres.shapes.RECTANGLE, { x:rx, y:cyy, w:rw, h:chh-0.06, fill:{color: i%2?"F2EEE7":"FFFFFF"}, line:{color:LIND,width:0.5} });
      s.addShape(pres.shapes.RECTANGLE, { x:rx, y:cyy, w:0.06, h:chh-0.06, fill:{color:RED} });
      s.addText(c[0], { x:rx+0.18, y:cyy, w:1.7, h:chh-0.06, margin:0, fontFace:BODY, bold:true, fontSize:10, color:INK, valign:"middle" });
      s.addText(c[1], { x:rx+1.95, y:cyy, w:rw-2.05, h:chh-0.06, margin:0, fontFace:BODY, fontSize:8.7, color:"55504A", valign:"middle" });
      cyy+=chh;
    });
    footer(s);
  }

  // ============================================================
  // SLIDE 12 — 메자닌 DSCR 민감도 (EBITDA 상향)
  // ============================================================
  {
    const s = pres.addSlide(); base(s);
    header(s, "FINANCING · MEZZANINE DSCR SENSITIVITY", "12 / 12",
      [{ text:"메자닌 $1bn 유치 조건 — ex-AMPC EBITDA를 얼마나 올려야 하나", options:{} }],
      "메자닌 $1,000mm · 배당 7% · AMPC창내(2029~32) 거치상환 · DSCR = CFADS ÷ (DOE이자+메자닌 원리금)");

    // 민감도 표
    const ty=2.7, tW=W-2*MX;
    const rows=[
      ["5%  (DC블록 박리다매)","1.36","1.38","1.13","0.70","✗ 2031~32 미달",false],
      ["7%","1.45","1.49","1.26","0.84","✗ 2031~32 미달",false],
      ["9%  ★ 채택","1.53","1.60","1.39","0.98","✓ 핵심 3개년 충족",true],
      ["11%","1.61","1.71","1.52","1.12","✓ 2032 경계",false],
      ["13%","1.70","1.82","1.65","1.26","✓ 거의 전구간",false],
      ["15%","1.78","1.92","1.78","1.40","✓ 전구간 충족",false],
    ];
    // header
    const cols=["ex-AMPC EBITDA","2029","2030","2031","2032","판정"];
    const cw=[3.0,1.4,1.4,1.4,1.4, tW-8.6];
    let cx=MX;
    s.addShape(pres.shapes.RECTANGLE, { x:MX, y:ty, w:tW, h:0.45, fill:{color:INK} });
    cols.forEach((c,i)=>{
      s.addText(c, { x:cx+0.1, y:ty, w:cw[i]-0.15, h:0.45, margin:0, align: i===0?"left":(i===5?"left":"center"), fontFace:BODY, bold:true, color:WHITE, fontSize:11, valign:"middle" });
      cx+=cw[i];
    });
    let yy=ty+0.45;
    rows.forEach(rw=>{
      const hl=rw[6];
      s.addShape(pres.shapes.RECTANGLE, { x:MX, y:yy, w:tW, h:0.56, fill:{color: hl?"FBF3F2":"FFFFFF"}, line:{color: hl?RED:LIND, width: hl?1.5:0.5} });
      cx=MX;
      // EBITDA 라벨
      s.addText(rw[0], { x:cx+0.12, y:yy, w:cw[0]-0.15, h:0.56, margin:0, fontFace:BODY, bold:hl, fontSize:11, color: hl?RED:INK, valign:"middle" });
      cx+=cw[0];
      // DSCR 4개
      for(let i=1;i<=4;i++){
        const val=parseFloat(rw[i]); const ok=val>=1.3;
        s.addText(rw[i]+"x", { x:cx, y:yy, w:cw[i], h:0.56, margin:0, align:"center", fontFace:MONO, bold:true, fontSize:11.5, color: ok?POS:(val>=1.0?MUTE:RED), valign:"middle" });
        cx+=cw[i];
      }
      // 판정
      s.addText(rw[5], { x:cx+0.1, y:yy, w:cw[5]-0.15, h:0.56, margin:0, fontFace:BODY, fontSize:9.5, bold:hl, color: hl?RED:"55504A", valign:"middle" });
      yy+=0.56;
    });
    s.addText("DSCR ≥ 1.3x = 메자닌 적격 (녹색) · 1.0~1.3x = 경계(회색) · <1.0x = 미달(적색)", { x:MX, y:yy+0.08, w:tW, h:0.3, margin:0, fontFace:BODY, fontSize:9, color:MUTE, italic:true });

    // 결론 박스 2개
    const by=yy+0.55, bw=(tW-0.4)/2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:MX, y:by, w:bw, h:1.5, rectRadius:0.05, fill:{color:"F2ECE0"}, line:{color:CMPD,width:1} });
    s.addText("결론 — 9%가 현실적 타겟", { x:MX+0.2, y:by+0.13, w:bw-0.4, h:0.34, margin:0, fontFace:HEAD, bold:true, fontSize:13, color:INK });
    s.addText([
      {text:"• ex-AMPC EBITDA 9%면 핵심 3개년(2029~31) DSCR 1.39~1.60x로 메자닌 기준 충족\n",options:{breakLine:true}},
      {text:"• 2032(AMPC 거의 소진)는 만기 1년 단축 또는 SK 잔액 흡수로 처리\n",options:{breakLine:true}},
      {text:"• 5%→9% 상향 = DC블록을 박리다매가 아닌 통합마진 시스템으로 판매",options:{}},
    ], { x:MX+0.2, y:by+0.5, w:bw-0.4, h:0.95, margin:0, fontFace:BODY, fontSize:9.5, color:"55504A", valign:"top", lineSpacingMultiple:1.05 });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:MX+bw+0.4, y:by, w:bw, h:1.5, rectRadius:0.05, fill:{color:INK} });
    s.addText("두 가지 해법 — 트레이드오프", { x:MX+bw+0.6, y:by+0.13, w:bw-0.4, h:0.34, margin:0, fontFace:HEAD, bold:true, fontSize:13, color:WHITE });
    s.addText([
      {text:"방법1: ",options:{bold:true,color:"FF9B9B"}},{text:"EBITDA 5% 유지 + 메자닌 $700mm 축소\n",options:{color:WHITE,breakLine:true}},
      {text:"방법2: ",options:{bold:true,color:"9BD3A8"}},{text:"메자닌 $1bn 유지 + EBITDA 9% 상향 (채택)\n",options:{color:WHITE,breakLine:true}},
      {text:"→ '$300mm 추가자금이 절실한가' vs 'EBITDA 4%p 개선이 가능한가'의 선택",options:{color:"C9C1B2",italic:true}},
    ], { x:MX+bw+0.6, y:by+0.5, w:bw-0.4, h:0.95, margin:0, fontFace:BODY, fontSize:9.5, valign:"top", lineSpacingMultiple:1.05 });

    s.addText("EV 효과: ex-AMPC EBITDA 5%→9% 상향으로 DCF EV $1.83bn → $3.16bn (마진 개선이 가치에 직결).",
      { x:MX, y:Hh-0.62, w:W-2*MX, h:0.3, margin:0, fontFace:BODY, fontSize:9.5, color:RED, italic:true, valign:"top" });
    footer(s);
  }

  await pres.writeFile({ fileName: require("path").join(__dirname,"..","output","SK_Valuation_Financing.pptx") });
  console.log("WROTE deck");
})();
