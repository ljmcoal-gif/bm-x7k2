// SK Valuation Deck v2 — SK 실적발표 템플릿 디자인 · PPT A4 가로
// 디자인 규칙: 배경 #FBF9F5 / 헤드라인 Noto Serif KR Bold / 본문 IBM Plex Sans KR
//             수치 IBM Plex Mono / 레드 #C1121F / 웜그레이 #D8CFC2 / 그린 #1E7A4D
//             그라데이션·3D·그림자 금지 · 가로 그리드만
const pptxgen = require(require('child_process').execSync('npm root -g').toString().trim() + '/pptxgenjs');
const pres = new pptxgen();

// ===== A4 가로 =====
const W = 11.69, H = 8.27;
pres.defineLayout({ name: 'A4L', width: W, height: H });
pres.layout = 'A4L';

// ===== 디자인 토큰 =====
const BG   = "FBF9F5";   // 페이퍼 배경
const INK  = "1B1714";   // 텍스트/기준선
const DARK = "1B1714";   // 다크 섹션 배경
const RED  = "C1121F";   // 강조/당분기/핵심
const WARM = "D8CFC2";   // 비교/과거 계열
const GREEN= "1E7A4D";   // 증감 양수
const GRID = "ECE6DC";   // 가로 그리드선
const RULE = "E4DED4";   // 옅은 구분선
const MUTE = "6F6960";   // 보조 텍스트
const MUTE2= "9A847C";   // 더 옅은 보조(모노 라벨)
const PAPER= "FFFFFF";

// 폰트
const HEAD = "Noto Serif CJK KR";   // 헤드라인 명조
const BODY = "IBM Plex Sans KR";    // 본문
const MONO = "IBM Plex Mono";       // 수치·라벨

// 여백
const MX = 0.62;        // 좌우 마진
const TOP = 0.52;       // 상단 eyebrow y

// ===== 공통 헬퍼 =====
function base(s){ s.background = { color: BG }; }

// 상단 eyebrow + 헤드라인 + 우측 페이지
function header(s, eyebrow, page, title, dek){
  // eyebrow (모노, 레드 좌 / 회색 우)
  s.addText(eyebrow, { x:MX, y:TOP, w:7.5, h:0.26, margin:0, align:"left",
    fontFace:MONO, fontSize:9, color:RED, charSpacing:2, valign:"middle" });
  s.addText("SK Innovation · IR  /  "+page, { x:W-4.5-MX, y:TOP, w:4.5, h:0.26, margin:0, align:"right",
    fontFace:MONO, fontSize:9, color:MUTE2, charSpacing:1.5, valign:"middle" });
  // 헤드라인 (명조 Bold) + 하단 2px 기준선
  s.addText(title, { x:MX, y:TOP+0.34, w:W-2*MX-2.4, h:0.62, margin:0, align:"left",
    fontFace:HEAD, fontSize:26, bold:true, color:INK, valign:"middle" });
  if(dek){
    s.addText(dek, { x:W-MX-3.0, y:TOP+0.5, w:3.0, h:0.42, margin:0, align:"right",
      fontFace:BODY, fontSize:9.5, color:MUTE, valign:"bottom" });
  }
  s.addShape(pres.shapes.LINE, { x:MX, y:TOP+1.04, w:W-2*MX, h:0,
    line:{ color:INK, width:1.5 } });
}

function footer(s, page){
  s.addShape(pres.shapes.LINE, { x:MX, y:H-0.46, w:W-2*MX, h:0, line:{ color:RULE, width:0.75 } });
  s.addText("Project [ ● ]  ·  Strictly Private & Confidential", { x:MX, y:H-0.42, w:5, h:0.26, margin:0,
    fontFace:BODY, fontSize:7.5, color:MUTE2, valign:"middle" });
  s.addText("DRAFT — FOR DISCUSSION ONLY", { x:W/2-2, y:H-0.42, w:4, h:0.26, margin:0, align:"center",
    fontFace:MONO, fontSize:7.5, color:MUTE2, charSpacing:1.5, valign:"middle" });
  s.addText("2026", { x:W-MX-1.5, y:H-0.42, w:1.5, h:0.26, margin:0, align:"right",
    fontFace:MONO, fontSize:7.5, color:MUTE2, valign:"middle" });
}

// 섹션 소제목 (모노 eyebrow)
function eyebrowLabel(s, x, y, w, txt){
  s.addText(txt, { x, y, w, h:0.24, margin:0, fontFace:MONO, fontSize:8.5, color:MUTE2,
    charSpacing:1.5, valign:"middle" });
}

// ============================================================
(async ()=>{

  // ════════════ SLIDE 1 — 표지 ════════════
  {
    const s = pres.addSlide(); base(s);
    // 배경 워터마크 큰 글자 (연한 레드, 명조)
    s.addText("EV", { x:W-4.6, y:0.5, w:4.4, h:4.4, margin:0, align:"right",
      fontFace:HEAD, fontSize:340, bold:true, color:RED, transparency:94, valign:"top" });
    // 상단
    s.addText([{text:"",options:{}}], {x:0,y:0,w:0.01,h:0.01}); // noop
    s.addShape(pres.shapes.RECTANGLE, { x:MX, y:0.62, w:0.42, h:0.055, fill:{color:RED} });
    s.addText("SK Innovation", { x:MX+0.52, y:0.5, w:4, h:0.3, margin:0, fontFace:BODY, bold:true, fontSize:13, color:INK, valign:"middle" });
    s.addText("INVESTOR RELATIONS", { x:W-MX-3.5, y:0.5, w:3.5, h:0.3, margin:0, align:"right",
      fontFace:MONO, fontSize:10, color:MUTE2, charSpacing:2.5, valign:"middle" });
    // 메인 타이틀
    s.addText("M&A VALUATION & FINANCING", { x:MX, y:H-3.5, w:8, h:0.34, margin:0,
      fontFace:MONO, fontSize:13, color:RED, charSpacing:3, valign:"middle" });
    s.addText("US LFP/ESS DC블록\n플랫폼 인수 검토", { x:MX, y:H-3.0, w:9, h:1.7, margin:0,
      fontFace:HEAD, fontSize:54, bold:true, color:INK, lineSpacingMultiple:1.05, valign:"top" });
    // 하단 메타
    s.addShape(pres.shapes.LINE, { x:MX, y:H-0.95, w:W-2*MX, h:0, line:{ color:RULE, width:1 } });
    s.addText("기업가치 · 파이낸싱 구조 · FCFE 분석", { x:MX, y:H-0.88, w:6, h:0.3, margin:0,
      fontFace:BODY, fontSize:11, color:MUTE, valign:"middle" });
    s.addText("SKBA (조지아) + SKOT (테네시)", { x:W-MX-5, y:H-0.88, w:5, h:0.3, margin:0, align:"right",
      fontFace:MONO, fontSize:11, color:MUTE, valign:"middle" });
  }

  // ════════════ SLIDE 2 — 풋볼필드 ════════════
  {
    const s = pres.addSlide(); base(s);
    header(s, "VALUATION SUMMARY — FOOTBALL FIELD", "02",
      "통합 플랫폼 기업가치 — 3개 방법론 수렴",
      "DCF $3.5bn · 멀티플 · NAV $5.0bn");

    // chart geometry
    const cx = MX, cy = 2.0, chartW = 7.7, chartH = 4.4;
    const axX = cx + 2.55, axW = chartW - 2.55 - 0.1;
    const v2x = v => axX + (v/7.0)*axW;
    const gTop = cy+0.15, gBot = cy+chartH-0.5;

    eyebrowLabel(s, cx, cy-0.18, 5, "ENTERPRISE VALUE  ($ bn)");

    // 결론 밴드 3.0-4.0
    s.addShape(pres.shapes.RECTANGLE, { x:v2x(3.0), y:gTop, w:v2x(4.0)-v2x(3.0), h:gBot-gTop,
      fill:{ color:RED, transparency:90 }, line:{type:"none"} });
    // 세로 그리드 (옅게) + 축 라벨
    [0,1,2,3,4,5,6,7].forEach(t=>{
      const major = (t===0);
      s.addShape(pres.shapes.LINE, { x:v2x(t), y:gTop, w:0, h:gBot-gTop,
        line:{ color: major?INK:GRID, width: major?1.5:0.75 } });
      s.addText(String(t), { x:v2x(t)-0.3, y:gBot+0.04, w:0.6, h:0.24, margin:0, align:"center",
        fontFace:MONO, fontSize:9, color:MUTE2, valign:"middle" });
    });
    // 결론선 @3.45
    s.addShape(pres.shapes.LINE, { x:v2x(3.45), y:gTop-0.2, w:0, h:(gBot-gTop)+0.2, line:{ color:RED, width:1.5, dashType:"dash" } });
    s.addText("결론 EV ≈ $3.5bn", { x:v2x(3.45)+0.05, y:gTop-0.22, w:2.0, h:0.26, margin:0.02,
      fill:{color:RED}, color:PAPER, fontFace:BODY, bold:true, fontSize:9, align:"center", valign:"middle" });

    // rows
    const rows = [
      ["DCF (DC블록·9%·NOL)", "WACC 9~12% · TGR 2%", 2.97, 4.11, true],
      ["EV/EBITDA ex-AMPC (2029E)", "fwd 8–14x · $350mm", 2.80, 4.90, false],
      ["EV/EBITDA ex-AMPC (run-rate)", "fwd 8–14x · $441mm", 3.53, 6.17, false],
      ["순자산 NAV (SOTP)", "SKBA $2.0 + SKOT $3.0", 5.00, 5.00, false],
    ];
    const rowH=(gBot-gTop)/rows.length;
    rows.forEach((r,i)=>{
      const ry=gTop+i*rowH+rowH/2;
      // 라벨
      s.addText(r[0], { x:cx, y:ry-0.26, w:2.5, h:0.28, margin:0, fontFace:BODY, bold:true, fontSize:9.5, color:INK, valign:"middle" });
      s.addText(r[1], { x:cx, y:ry+0.0, w:2.5, h:0.22, margin:0, fontFace:MONO, fontSize:7.5, color:MUTE2, valign:"middle" });
      const bx=v2x(r[2]), bw=v2x(r[3])-v2x(r[2]);
      const pt = Math.abs(r[3]-r[2])<0.01;
      if(pt){
        s.addShape(pres.shapes.DIAMOND, { x:bx-0.11, y:ry-0.13, w:0.26, h:0.26, fill:{color:INK}, line:{type:"none"} });
        s.addText("$"+r[2].toFixed(1)+"bn", { x:bx+0.2, y:ry-0.13, w:1.1, h:0.26, margin:0, align:"left", fontFace:MONO, fontSize:9, bold:true, color:INK, valign:"middle" });
      } else {
        s.addShape(pres.shapes.RECTANGLE, { x:bx, y:ry-0.13, w:bw, h:0.26, fill:{ color: r[4]?RED:WARM }, line:{type:"none"} });
        s.addText(r[2].toFixed(2), { x:bx-0.56, y:ry-0.13, w:0.5, h:0.26, margin:0, align:"right", fontFace:MONO, fontSize:9, bold:true, color:INK, valign:"middle" });
        s.addText(r[3].toFixed(2), { x:bx+bw+0.06, y:ry-0.13, w:0.5, h:0.26, margin:0, align:"left", fontFace:MONO, fontSize:9, bold:true, color:INK, valign:"middle" });
      }
    });

    // 우측 NAV 카드
    const nx=cx+chartW+0.35, nw=W-nx-MX;
    eyebrowLabel(s, nx, cy-0.18, nw, "순자산가치 — 엔티티별 (SOTP)");
    const cards=[
      ["SKBA — 조지아","가동중",[["순자산 NAV","~$2.0 bn ★"],["캐파(파우치)","21.2 GWh"],["개조 capex","$0.4 bn"],["COD","'27 / '28"]]],
      ["SKOT — 테네시","'28 가동",[["순자산 NAV","$3.0 bn"],["캐파","29.7 GWh"],["DOE loan","$4.0 bn"],["메자닌","$1.0 bn"]]],
    ];
    let ny=cy+0.12;
    cards.forEach(c=>{
      const ch=1.95;
      s.addShape(pres.shapes.RECTANGLE, { x:nx, y:ny, w:nw, h:ch, fill:{color:PAPER}, line:{color:RULE,width:1} });
      s.addShape(pres.shapes.RECTANGLE, { x:nx, y:ny, w:0.05, h:ch, fill:{color:RED} });
      s.addText(c[0], { x:nx+0.18, y:ny+0.1, w:nw-0.6, h:0.28, margin:0, fontFace:BODY, bold:true, fontSize:11, color:INK, valign:"middle" });
      s.addText(c[1], { x:nx+nw-0.95, y:ny+0.1, w:0.85, h:0.28, margin:0, align:"right", fontFace:MONO, fontSize:8, color:MUTE2, valign:"middle" });
      let ky=ny+0.46;
      c[2].forEach(kv=>{
        s.addText(kv[0], { x:nx+0.18, y:ky, w:nw*0.5, h:0.3, margin:0, fontFace:BODY, fontSize:9, color:MUTE, valign:"middle" });
        s.addText(kv[1], { x:nx+nw*0.42, y:ky, w:nw*0.58-0.18, h:0.3, margin:0, align:"right", fontFace:MONO, fontSize:9.5, bold:true, color: kv[1].includes('★')?RED:INK, valign:"middle" });
        s.addShape(pres.shapes.LINE, { x:nx+0.18, y:ky+0.3, w:nw-0.36, h:0, line:{color:GRID,width:0.5} });
        ky+=0.33;
      });
      ny+=ch+0.18;
    });

    footer(s, "02");
  }

  // ════════════ SLIDE 3 — 밸류 브리지 (AMPC 의존도) ════════════
  {
    const s = pres.addSlide(); base(s);
    header(s, "VALUE BRIDGE — AMPC 의존도", "03",
      "기업가치의 원천 — 보조금과 사업 본연",
      "AMPC는 2033 소멸");

    eyebrowLabel(s, MX, 1.95, 7, "EBITDA 구성  ·  $mm  ·  사업 본연(레드) + AMPC(웜그레이)");

    const yrs=["2027","2028","2029","2030","2031","2032","2033"];
    const ebXa  =[72,262,350,414,448,459,453];   // 사업 본연(ex-AMPC)
    const ampc  =[210,851,1160,1048,767,399,0];  // AMPC

    const chX=MX, chY=2.4, chW=7.5, chH=4.25;
    const maxV=1600;
    const slot=(chW-0.55)/yrs.length;
    const bw=slot*0.52;
    const y0=chY+chH-0.35;
    const v2y=(v)=> y0 - (v/maxV)*(chH-0.5);
    // 가로 그리드
    [0,400,800,1200,1600].forEach(g=>{
      const gy=v2y(g);
      s.addShape(pres.shapes.LINE, { x:chX+0.55, y:gy, w:chW-0.55, h:0, line:{ color: g===0?INK:GRID, width: g===0?2:1 } });
      s.addText("$"+g, { x:chX, y:gy-0.1, w:0.5, h:0.2, margin:0, align:"right", fontFace:MONO, fontSize:7.5, color:MUTE2, valign:"middle" });
    });
    // 적층 막대
    yrs.forEach((yr,i)=>{
      const bx=chX+0.55+i*slot+(slot-bw)/2;
      const xaH=y0-v2y(ebXa[i]);
      const amH=y0-v2y(ebXa[i]+ampc[i]) - xaH;
      // 사업 본연 (레드, 아래)
      s.addShape(pres.shapes.RECTANGLE, { x:bx, y:y0-xaH, w:bw, h:xaH, fill:{color:RED}, line:{type:"none"} });
      // AMPC (웜그레이, 위)
      if(ampc[i]>0) s.addShape(pres.shapes.RECTANGLE, { x:bx, y:y0-xaH-amH, w:bw, h:amH, fill:{color:WARM}, line:{type:"none"} });
      // 합계 라벨
      s.addText("$"+(ebXa[i]+ampc[i]), { x:bx-0.15, y:y0-xaH-amH-0.22, w:bw+0.3, h:0.2, margin:0, align:"center", fontFace:MONO, fontSize:7.5, bold:true, color:INK, valign:"middle" });
      // 연도
      s.addText(yr, { x:bx-0.15, y:y0+0.05, w:bw+0.3, h:0.2, margin:0, align:"center", fontFace:MONO, fontSize:8, color: i===2?RED:MUTE2, valign:"middle" });
    });
    // 범례
    s.addShape(pres.shapes.RECTANGLE, { x:chX+0.55, y:chY+chH+0.05, w:0.18, h:0.13, fill:{color:RED}, line:{type:"none"} });
    s.addText("사업 본연 (ex-AMPC EBITDA)", { x:chX+0.78, y:chY+chH-0.02, w:2.6, h:0.22, margin:0, fontFace:BODY, fontSize:8, color:INK, valign:"middle" });
    s.addShape(pres.shapes.RECTANGLE, { x:chX+3.5, y:chY+chH+0.05, w:0.18, h:0.13, fill:{color:WARM}, line:{type:"none"} });
    s.addText("AMPC 45X (2033 소멸)", { x:chX+3.73, y:chY+chH-0.02, w:2.6, h:0.22, margin:0, fontFace:BODY, fontSize:8, color:INK, valign:"middle" });

    // 우측 인사이트
    const rx=chX+chW+0.35, rw=W-rx-MX;
    eyebrowLabel(s, rx, 1.95, rw, "KEY INSIGHT");
    const ins=[
      ["76%","2029년 EBITDA 중 AMPC 비중 — 정책 의존 정점"],
      ["$453mm","2033 AMPC 소멸 후에도 사업 본연 EBITDA 유지(마진 9%)"],
      ["8–11x","ex-AMPC EBITDA 멀티플이 DCF $3.5bn과 수렴"],
    ];
    let iy=2.3;
    ins.forEach(it=>{
      s.addShape(pres.shapes.RECTANGLE, { x:rx, y:iy, w:rw, h:1.25, fill:{color:PAPER}, line:{color:RULE,width:1} });
      s.addShape(pres.shapes.RECTANGLE, { x:rx, y:iy, w:0.05, h:1.25, fill:{color:RED} });
      s.addText(it[0], { x:rx+0.2, y:iy+0.14, w:rw-0.4, h:0.5, margin:0, fontFace:MONO, bold:true, fontSize:26, color:RED, valign:"middle" });
      s.addText(it[1], { x:rx+0.2, y:iy+0.66, w:rw-0.4, h:0.5, margin:0, fontFace:BODY, fontSize:9, color:MUTE, valign:"top" });
      iy+=1.42;
    });

    footer(s, "03");
  }

  // ════════════ SLIDE 4 — Peer 멀티플 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"COMPARABLE COMPANIES — TRADING MULTIPLES","04",
      "배터리 셀 & ESS Peer — 트레이딩 멀티플","시장데이터 2026.6 · indicative");
    const cols=["Company","EV/Rev '26E","EV/EBITDA '26E","P/B"];
    const cw=[4.6,1.85,2.0,1.6];
    let hx=MX;
    s.addShape(pres.shapes.LINE,{x:MX,y:1.95,w:W-2*MX,h:0,line:{color:INK,width:1.5}});
    cols.forEach((c,i)=>{ s.addText(c,{x:hx+0.1,y:2.0,w:cw[i]-0.15,h:0.32,margin:0,align:i===0?"left":"center",fontFace:MONO,fontSize:9,color:MUTE2,charSpacing:1,valign:"middle"}); hx+=cw[i]; });
    const rows=[
      ["대형 셀 제조사",null,null,null,true],
      ["LG Energy Solution","3.8x","11x","2.4x",false],
      ["Samsung SDI","1.6x","8.6x","1.1x",false],
      ["CATL","3.0x","13x","4.6x",false],
      ["미국/ESS 시스템·첨단셀",null,null,null,true],
      ["Fluence Energy","1.2x","37x*","5.1x",false],
      ["Tesla — Energy","≈2.0x","≈8x","—",false],
      ["EVE Energy","2.0x","13x","3.0x",false],
      ["PEER MEDIAN","2.0x","8–13x","3.0x",false,false,true],
    ];
    let ry=2.4;
    rows.forEach(r=>{
      if(r[4]){ // 섹션
        s.addText("— "+r[0]+" —",{x:MX,y:ry,w:6,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:9,color:RED,valign:"middle"}); ry+=0.34; return;
      }
      const med=r[5];
      if(med) s.addShape(pres.shapes.LINE,{x:MX,y:ry,w:W-2*MX,h:0,line:{color:INK,width:1}});
      hx=MX;
      [r[0],r[1],r[2],r[3]].forEach((c,i)=>{
        s.addText(c||"",{x:hx+0.1,y:ry,w:cw[i]-0.15,h:0.4,margin:0,align:i===0?"left":"center",
          fontFace:i===0?BODY:MONO,bold:(i===0||med),fontSize:i===0?9.5:10.5,color:med?RED:INK,valign:"middle"});
        hx+=cw[i];
      });
      s.addShape(pres.shapes.LINE,{x:MX,y:ry+0.44,w:W-2*MX,h:0,line:{color:GRID,width:0.5}});
      ry+=0.46;
    });
    // implied 박스
    ry+=0.15;
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:ry,w:W-2*MX,h:0.42,fill:{color:DARK},line:{type:"none"}});
    s.addText("우리 회사 적용 — EV/EBITDA ex-AMPC (fwd)",{x:MX+0.15,y:ry,w:5,h:0.42,margin:0,fontFace:BODY,bold:true,fontSize:10,color:PAPER,valign:"middle"});
    s.addText("Fwd EBITDA ex-AMPC $350M(2029E) / run-rate $441M",{x:W-MX-5.2,y:ry,w:5.05,h:0.42,margin:0,align:"right",fontFace:MONO,fontSize:8,color:WARM,valign:"middle"});
    ry+=0.42;
    const iv=[["8x","$2.8bn"],["11x","$3.9bn"],["14x","$4.9bn"],["DCF","$3.5bn"]];
    const ivw=(W-2*MX)/4;
    iv.forEach((v,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:MX+i*ivw,y:ry,w:ivw,h:0.62,fill:{color: i===3?"F2ECE0":PAPER},line:{color:RULE,width:0.5}});
      s.addText(v[0],{x:MX+i*ivw+0.12,y:ry+0.05,w:ivw-0.24,h:0.26,margin:0,fontFace:MONO,fontSize:8.5,color:MUTE2,valign:"middle"});
      s.addText(v[1],{x:MX+i*ivw+0.12,y:ry+0.28,w:ivw-0.24,h:0.3,margin:0,fontFace:MONO,bold:true,fontSize:13,color: i===3?RED:INK,valign:"middle"});
    });
    s.addText("* Fluence n.m. 제외 · peer EBITDA에도 AMPC 포함 → 보수적 비교",{x:MX,y:H-0.66,w:8,h:0.2,margin:0,fontFace:BODY,fontSize:8,color:MUTE2,valign:"middle"});
    footer(s,"04");
  }

  // ════════════ SLIDE 5 — DCF 가정 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"DCF & PROJECTION ASSUMPTIONS","05",
      "주요 DCF · 프로젝션 가정","Base case · ★ 미검증");
    const colW=(W-2*MX-0.5)/2;
    // 좌: 운영
    eyebrowLabel(s,MX,1.95,colW,"운영 · 매출 가정");
    const op=[["제품/시장","LFP DC블록 · ESS"],["DC블록 ASP","$155 /kWh"],["ex-AMPC EBITDA","9.0%"],["AMPC 단가","$45 → 0 ('33)"],["명판 — SKBA","21.2 GWh"],["명판 — SKOT","29.7 GWh"],["개조 capex","$1.4 bn"]];
    let ly=2.3;
    op.forEach(kv=>{
      s.addText(kv[0],{x:MX,y:ly,w:colW*0.55,h:0.36,margin:0,fontFace:BODY,fontSize:10,color:MUTE,valign:"middle"});
      s.addText(kv[1],{x:MX+colW*0.4,y:ly,w:colW*0.6,h:0.36,margin:0,align:"right",fontFace:MONO,bold:true,fontSize:10.5,color:INK,valign:"middle"});
      s.addShape(pres.shapes.LINE,{x:MX,y:ly+0.36,w:colW,h:0,line:{color:GRID,width:0.5}}); ly+=0.4;
    });
    // 우: DCF 파라미터 + 민감도
    const rx=MX+colW+0.5;
    eyebrowLabel(s,rx,1.95,colW,"DCF 핵심 파라미터");
    const dcf=[["예측기간","2026–2035"],["WACC","9.0–12.0% (base 10.5%)"],["TGR","2.0%"],["Tax","23% (AMPC 비과세)"],["SKBA NOL","$1.5 bn (TCJA 80%)"],["순부채 (DOE)","$4.0 bn"]];
    let ry=2.3;
    dcf.forEach(kv=>{
      s.addText(kv[0],{x:rx,y:ry,w:colW*0.5,h:0.34,margin:0,fontFace:BODY,fontSize:10,color:MUTE,valign:"middle"});
      s.addText(kv[1],{x:rx+colW*0.35,y:ry,w:colW*0.65,h:0.34,margin:0,align:"right",fontFace:MONO,bold:true,fontSize:10,color:INK,valign:"middle"});
      s.addShape(pres.shapes.LINE,{x:rx,y:ry+0.34,w:colW,h:0,line:{color:GRID,width:0.5}}); ry+=0.38;
    });
    // 결론 EV
    s.addShape(pres.shapes.RECTANGLE,{x:rx,y:ry+0.05,w:colW,h:0.46,fill:{color:DARK},line:{type:"none"}});
    s.addText("결론 EV (base)",{x:rx+0.15,y:ry+0.05,w:colW*0.6,h:0.46,margin:0,fontFace:BODY,bold:true,fontSize:11,color:PAPER,valign:"middle"});
    s.addText("≈ $3.5 bn",{x:rx+colW*0.5,y:ry+0.05,w:colW*0.5-0.15,h:0.46,margin:0,align:"right",fontFace:MONO,bold:true,fontSize:13,color:PAPER,valign:"middle"});
    ry+=0.7;
    // 민감도 표
    eyebrowLabel(s,rx,ry,colW,"민감도 — EV($bn) · WACC × TGR"); ry+=0.3;
    const sens=[["WACC＼TGR","1.5%","2.0%","2.5%"],["9.0%","3.99","4.11","4.24"],["10.5%","3.38","3.45","3.53"],["12.0%","2.92","2.97","3.02"]];
    const scw=colW/4;
    sens.forEach((row,ri)=>{
      row.forEach((c,ci)=>{
        const hd=ri===0, rh=ci===0, mid=(ri===2&&ci===2);
        s.addShape(pres.shapes.RECTANGLE,{x:rx+ci*scw,y:ry,w:scw,h:0.34,fill:{color: hd?DARK : mid?"F2DCDE" : rh?"F2ECE0":PAPER},line:{color:RULE,width:0.5}});
        s.addText(c,{x:rx+ci*scw,y:ry,w:scw,h:0.34,margin:0,align:"center",fontFace:hd||!rh?MONO:BODY,bold:hd||rh||mid,fontSize:hd?8:9.5,color:hd?PAPER:(mid?RED:INK),valign:"middle"});
      });
      ry+=0.34;
    });
    s.addText("SKBA NOL $1.5bn(누적결손) 반영 → 초기 세금 절감으로 EV +$0.3bn",{x:MX,y:H-0.66,w:8,h:0.2,margin:0,fontFace:BODY,fontSize:8,color:RED,valign:"middle"});
    footer(s,"05");
  }

  // ════════════ SLIDE 6 — Crucible 비교 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"FINANCING — CRUCIBLE BENCHMARK","06",
      "고려아연 Crucible 대비 — 차용 vs 차별","정부 JV · 제3자배정");
    const colW=(W-2*MX-0.4)/2;
    // 좌 차용
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:1.95,w:colW,h:0.4,fill:{color:INK},line:{type:"none"}});
    s.addText("✓ 차용한 골격",{x:MX+0.15,y:1.95,w:colW-0.3,h:0.4,margin:0,fontFace:BODY,bold:true,fontSize:12,color:PAPER,valign:"middle"});
    const bor=[["정부 앵커 우선","DoE/DoD/DFC로 FAST-41·FEOC·CFIUS 옵틱 정리"],["정부·SI 90%+ 부담","SK 잔여출자 → 재무안정성↑·운영주도권 유지"],["HoldCo/SPV 다층","자금 성격별로 칸을 나눠 받는 골격"]];
    let by=2.5;
    bor.forEach(b=>{
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:by,w:colW,h:0.95,fill:{color:PAPER},line:{color:RULE,width:1}});
      s.addText(b[0],{x:MX+0.15,y:by+0.1,w:colW-0.3,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:11,color:INK,valign:"middle"});
      s.addText(b[1],{x:MX+0.15,y:by+0.42,w:colW-0.3,h:0.48,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:"top"});
      by+=1.07;
    });
    // 우 차별
    const rx=MX+colW+0.4;
    s.addShape(pres.shapes.RECTANGLE,{x:rx,y:1.95,w:colW,h:0.4,fill:{color:RED},line:{type:"none"}});
    s.addText("✗ 결정적으로 다른 점",{x:rx+0.15,y:1.95,w:colW-0.3,h:0.4,margin:0,fontFace:BODY,bold:true,fontSize:12,color:PAPER,valign:"middle"});
    const dif=[["제3자배정 위치","KZ: 모회사 10% → 거버넌스 반발 / SK: 자회사 ≤19.9%"],["링펜싱 추가","KZ: 신설 격리불요 / SK: 공장2 DOE $4bn ring-fence"],["보조금 성격","KZ: CHIPS 설비보조 직접출자 / SK: AMPC 생산세액 유동화"]];
    let dy=2.5;
    dif.forEach(d=>{
      s.addShape(pres.shapes.RECTANGLE,{x:rx,y:dy,w:colW,h:0.95,fill:{color:"FBF3F2"},line:{color:RED,width:1}});
      s.addText(d[0],{x:rx+0.15,y:dy+0.1,w:colW-0.3,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:11,color:RED,valign:"middle"});
      s.addText(d[1],{x:rx+0.15,y:dy+0.42,w:colW-0.3,h:0.48,margin:0,fontFace:BODY,fontSize:9,color:INK,valign:"top"});
      dy+=1.07;
    });
    s.addText("요약: Crucible 골격(정부앵커+HoldCo+운영주도권)을 차용하되, 외국 의결권을 자회사 ≤19.9%로 한정 + 링펜싱 추가한 변형.",{x:MX,y:H-0.72,w:W-2*MX,h:0.3,margin:0,fontFace:BODY,fontSize:9,color:INK,valign:"middle"});
    footer(s,"06");
  }

  // ════════════ SLIDE 7 — 링펜싱 계약 스택 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"RING-FENCING — CONTRACT STACK","07",
      "링펜싱 계약 스택 — 5개 층 + non-recourse","SKOT 격리");
    const layers=[
      ["①","법인 분리","Bankruptcy-remote SPV + 독립이사 → 모회사 일방 도산·자산이전 차단"],
      ["②","LLC 약정","Separateness + Restricted Payments(DSCR 충족 후 배당) + Permitted Debt 한정"],
      ["③","자금조달 계약","우선주/메자닌 무의결권(→CoC 회피) + 우선배당·청산우선권"],
      ["④","채권자 간 약정","Subordination + Waterfall(운영비→DOE→메자닌→SK) + Standstill"],
      ["⑤","담보계약","45X AMPC 수취권리 담보·유동화 → 메자닌 상환재원"],
    ];
    let ly=2.05; const lh=0.82, lw=W-2*MX;
    layers.forEach((L,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:ly,w:lw,h:lh-0.1,fill:{color: i%2?"F2ECE0":PAPER},line:{color:RULE,width:1}});
      s.addShape(pres.shapes.OVAL,{x:MX+0.16,y:ly+0.16,w:0.4,h:0.4,fill:{color:RED},line:{type:"none"}});
      s.addText(L[0],{x:MX+0.16,y:ly+0.16,w:0.4,h:0.4,margin:0,align:"center",fontFace:HEAD,bold:true,fontSize:15,color:PAPER,valign:"middle"});
      s.addText(L[1],{x:MX+0.7,y:ly,w:2.3,h:lh-0.1,margin:0,fontFace:BODY,bold:true,fontSize:12,color:INK,valign:"middle"});
      s.addText(L[2],{x:MX+3.0,y:ly,w:lw-3.2,h:lh-0.1,margin:0,fontFace:BODY,fontSize:9.5,color:MUTE,valign:"middle"});
      ly+=lh;
    });
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:ly+0.05,w:lw,h:0.5,fill:{color:DARK},line:{type:"none"}});
    s.addText([{text:"관통 핵심 = Non-recourse:  ",options:{bold:true,color:PAPER,fontSize:11}},{text:"채무는 SKOT SPV 자산에만 청구 · 모회사·타 SPV 무소구 — 격벽의 본질.",options:{color:WARM,fontSize:10}}],{x:MX+0.2,y:ly+0.05,w:lw-0.4,h:0.5,margin:0,fontFace:BODY,valign:"middle"});
    footer(s,"07");
  }

  // ════════════ SLIDE 8 — 하이브리드 최적구조 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"FINANCING — OPTIMAL STRUCTURE","08",
      "최적 구조 — HoldCo 우산 + 법인별 격벽","waiver 전제");
    const colW=(W-2*MX-0.5)*0.5;
    eyebrowLabel(s,MX,1.95,colW,"세 가지 구조 비교"); 
    const opts=[["(A) 완전 분리","각자 조달·OT 격리","전염 차단·waiver 불요","OT 조달비↑·효율↓",false],
      ["(B) 하이브리드 ★","HoldCo 통합 + SPV 격벽","비용↓+격벽 유지 (최적)","waiver 협상 필요",true],
      ["(C) 완전 통합","Crucible형 단일 HoldCo","자본효율 최고","담보오염·전염·exit경직",false]];
    let oy=2.3;
    opts.forEach(o=>{
      const hl=o[4];
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:oy,w:colW,h:1.15,fill:{color: hl?"FBF3F2":PAPER},line:{color: hl?RED:RULE,width: hl?1.5:1}});
      s.addText(o[0],{x:MX+0.15,y:oy+0.1,w:colW-0.3,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:11.5,color: hl?RED:INK,valign:"middle"});
      s.addText(o[1],{x:MX+0.15,y:oy+0.4,w:colW-0.3,h:0.25,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:"top"});
      s.addText([{text:"＋ ",options:{color:GREEN,bold:true}},{text:o[2],options:{color:INK}}],{x:MX+0.15,y:oy+0.65,w:colW-0.3,h:0.24,margin:0,fontFace:BODY,fontSize:8.5,valign:"top"});
      s.addText([{text:"− ",options:{color:RED,bold:true}},{text:o[3],options:{color:MUTE}}],{x:MX+0.15,y:oy+0.88,w:colW-0.3,h:0.24,margin:0,fontFace:BODY,fontSize:8.5,valign:"top"});
      oy+=1.27;
    });
    // 우: 3 이유
    const rx=MX+colW+0.5, rw=W-rx-MX;
    eyebrowLabel(s,rx,1.95,rw,"waiver 받아도 분리 유지가 유리한 이유");
    const reas=[["담보 오염 방지","DOE 담보는 OT에만 → BA는 DOE-clean → 조달 저렴"],["FEOC/AMPC 적격성","적격성은 법인별 판정 → 분리하면 각자 깨끗"],["Exit 유연성","BA만 매각/OT만 carve-out 가능 (옵션가치)"]];
    let ry=2.3;
    reas.forEach((r,i)=>{
      s.addShape(pres.shapes.RECTANGLE,{x:rx,y:ry,w:rw,h:1.15,fill:{color:PAPER},line:{color:RULE,width:1}});
      s.addShape(pres.shapes.RECTANGLE,{x:rx,y:ry,w:0.05,h:1.15,fill:{color:RED}});
      s.addText((i+1)+". "+r[0],{x:rx+0.2,y:ry+0.13,w:rw-0.4,h:0.4,margin:0,fontFace:BODY,bold:true,fontSize:12,color:INK,valign:"middle"});
      s.addText(r[1],{x:rx+0.2,y:ry+0.55,w:rw-0.4,h:0.5,margin:0,fontFace:BODY,fontSize:9.5,color:MUTE,valign:"top"});
      ry+=1.27;
    });
    s.addText("waiver의 최적 용도 = 완전통합의 열쇠가 아니라, HoldCo→OT 비담보 후순위만 흘려보내는 분리 구조의 미세조정.",{x:MX,y:H-0.7,w:W-2*MX,h:0.3,margin:0,fontFace:BODY,fontSize:9,color:INK,valign:"middle"});
    footer(s,"08");
  }

  // ════════════ SLIDE 9 — 공장별 Financial Projection ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"FINANCIAL PROJECTION — BY ENTITY","09",
      "공장별 재무 프로젝션 — SKBA · SKOT · SK FCFE","$mm · ex-AMPC 9%·NOL");
    const yrs=["2027","2028","2029","2030","2031","2032","2033"];
    const D={ba_rev:[723,1643,1946,2177,2249,2244,2213],ba_ebit:[47,122,148,169,175,174,171],ba_ampc:[210,477,572,487,340,172,0],ba_fcff:[-30,394,638,576,456,299,130],ba_fcfe:[-76,348,593,530,410,254,84],
      ot_rev:[0,1242,1927,2420,2729,2861,2822],ot_ebit:[-47,45,107,151,179,191,187],ot_ampc:[0,374,588,561,428,227,0],ot_fcff:[-500,-40,582,598,504,335,132],ot_fcfe:[-723,-262,109,142,66,-85,-21],
      sk_fcfe:[-798,86,702,673,477,169,63]};
    function tbl(x,y,w,title,tint,rws){
      s.addShape(pres.shapes.RECTANGLE,{x,y,w,h:0.34,fill:{color:tint},line:{type:"none"}});
      s.addText(title,{x:x+0.12,y,w:w-0.24,h:0.34,margin:0,fontFace:BODY,bold:true,fontSize:10,color:PAPER,valign:"middle"});
      let yy=y+0.34; const lblW=1.7,cw=(w-lblW)/7;
      s.addShape(pres.shapes.RECTANGLE,{x,y:yy,w,h:0.26,fill:{color:"F2ECE0"},line:{type:"none"}});
      s.addText("($mm)",{x:x+0.1,y:yy,w:lblW-0.1,h:0.26,margin:0,fontFace:MONO,fontSize:7,color:MUTE2,valign:"middle"});
      yrs.forEach((yr,i)=>s.addText(yr,{x:x+lblW+i*cw,y:yy,w:cw,h:0.26,margin:0,align:"center",fontFace:MONO,bold:true,fontSize:7.5,color: i===2?RED:INK,valign:"middle"}));
      yy+=0.26;
      rws.forEach((rw,ri)=>{
        const vals=D[rw[1]],bd=rw[2];
        s.addShape(pres.shapes.RECTANGLE,{x,y:yy,w,h:0.28,fill:{color: ri%2?BG:PAPER},line:{color:GRID,width:0.5}});
        s.addText(rw[0],{x:x+0.1,y:yy,w:lblW-0.1,h:0.28,margin:0,fontFace:BODY,bold:bd,fontSize:8,color:INK,valign:"middle"});
        vals.forEach((v,i)=>{const ng=v<0; s.addText(ng?"("+Math.abs(v)+")":String(v),{x:x+lblW+i*cw,y:yy,w:cw,h:0.28,margin:0,align:"center",fontFace:MONO,bold:bd,fontSize:7.5,color:ng?RED:(bd?INK:MUTE),valign:"middle"});});
        yy+=0.28;
      });
      return yy;
    }
    const tW=W-2*MX; let cy2=2.0;
    cy2=tbl(MX,cy2,tW,"공장 1 — SKBA (조지아) · 파우치 21.2GWh",INK,[["매출","ba_rev",false],["EBIT (pre-AMPC)","ba_ebit",false],["AMPC 45X","ba_ampc",false],["FCFF","ba_fcff",true],["SKBA FCFE","ba_fcfe",true]]);
    cy2+=0.18;
    cy2=tbl(MX,cy2,tW,"공장 2 — SKOT (테네시) · 각형+파우치 29.7GWh",RED,[["매출","ot_rev",false],["EBIT (pre-AMPC)","ot_ebit",false],["AMPC 45X","ot_ampc",false],["FCFF","ot_fcff",true],["SKOT FCFE","ot_fcfe",true]]);
    cy2+=0.18;
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:cy2,w:tW,h:0.4,fill:{color:DARK},line:{type:"none"}});
    const lblW=1.7,cw=(tW-lblW)/7;
    s.addText("보통주(SK) FCFE 합계",{x:MX+0.1,y:cy2,w:lblW-0.1,h:0.4,margin:0,fontFace:BODY,bold:true,fontSize:9,color:PAPER,valign:"middle"});
    D.sk_fcfe.forEach((v,i)=>{const ng=v<0; s.addText(ng?"("+Math.abs(v)+")":String(v),{x:MX+lblW+i*cw,y:cy2,w:cw,h:0.4,margin:0,align:"center",fontFace:MONO,bold:true,fontSize:9,color:ng?"FF9B9B":PAPER,valign:"middle"});});
    footer(s,"09");
  }

  // ════════════ SLIDE 10 — 메자닌 DSCR 민감도 ════════════
  {
    const s=pres.addSlide(); base(s);
    header(s,"MEZZANINE DSCR SENSITIVITY","10",
      "메자닌 $1bn 유치 조건 — ex-AMPC EBITDA","DSCR=CFADS÷상환");
    const rows=[["5% (박리다매)","1.36","1.38","1.13","0.70","✗ 미달",false],
      ["7%","1.45","1.49","1.26","0.84","✗ 미달",false],
      ["9% ★ 채택","1.53","1.60","1.39","0.98","✓ 핵심3개년",true],
      ["11%","1.61","1.71","1.52","1.12","✓ 2032경계",false],
      ["13%","1.70","1.82","1.65","1.26","✓ 거의전구간",false],
      ["15%","1.78","1.92","1.78","1.40","✓ 전구간",false]];
    const cols=["ex-AMPC EBITDA","2029","2030","2031","2032","판정"];
    const tW=W-2*MX;
    const cw=[2.7,1.3,1.3,1.3,1.3, tW-2.7-1.3*4];
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:2.0,w:tW,h:0.4,fill:{color:DARK},line:{type:"none"}});
    let cx=MX;
    cols.forEach((c,i)=>{s.addText(c,{x:cx+0.1,y:2.0,w:cw[i]-0.15,h:0.4,margin:0,align:i===0||i===5?"left":"center",fontFace:BODY,bold:true,fontSize:10,color:PAPER,valign:"middle"}); cx+=cw[i];});
    let yy=2.4;
    rows.forEach(r=>{
      const hl=r[6];
      s.addShape(pres.shapes.RECTANGLE,{x:MX,y:yy,w:tW,h:0.5,fill:{color: hl?"FBF3F2":PAPER},line:{color: hl?RED:RULE,width: hl?1.5:0.5}});
      cx=MX;
      s.addText(r[0],{x:cx+0.12,y:yy,w:cw[0]-0.15,h:0.5,margin:0,fontFace:BODY,bold:hl,fontSize:10.5,color: hl?RED:INK,valign:"middle"}); cx+=cw[0];
      for(let i=1;i<=4;i++){const v=parseFloat(r[i]),ok=v>=1.3; s.addText(r[i]+"x",{x:cx,y:yy,w:cw[i],h:0.5,margin:0,align:"center",fontFace:MONO,bold:true,fontSize:11,color:ok?GREEN:(v>=1.0?MUTE:RED),valign:"middle"}); cx+=cw[i];}
      s.addText(r[5],{x:cx+0.1,y:yy,w:cw[5]-0.15,h:0.5,margin:0,fontFace:BODY,fontSize:9,bold:hl,color: hl?RED:MUTE,valign:"middle"});
      yy+=0.5;
    });
    yy+=0.2;
    const bw=(tW-0.4)/2;
    s.addShape(pres.shapes.RECTANGLE,{x:MX,y:yy,w:bw,h:1.3,fill:{color:"F2ECE0"},line:{color:RULE,width:1}});
    s.addText("결론 — 9%가 현실적 타겟",{x:MX+0.18,y:yy+0.12,w:bw-0.36,h:0.32,margin:0,fontFace:HEAD,bold:true,fontSize:12,color:INK,valign:"middle"});
    s.addText("핵심 3개년(2029~31) DSCR 1.39~1.60x 충족 · 2032는 만기단축/SK 흡수 · 5%→9%는 DC블록을 통합마진으로 판매",{x:MX+0.18,y:yy+0.5,w:bw-0.36,h:0.75,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:"top"});
    s.addShape(pres.shapes.RECTANGLE,{x:MX+bw+0.4,y:yy,w:bw,h:1.3,fill:{color:DARK},line:{type:"none"}});
    s.addText("두 가지 해법",{x:MX+bw+0.58,y:yy+0.12,w:bw-0.36,h:0.32,margin:0,fontFace:HEAD,bold:true,fontSize:12,color:PAPER,valign:"middle"});
    s.addText([{text:"방법1: ",options:{bold:true,color:"FF9B9B"}},{text:"EBITDA 5% + 메자닌 $700mm 축소\n",options:{color:PAPER}},{text:"방법2: ",options:{bold:true,color:"9BD3A8"}},{text:"메자닌 $1bn + EBITDA 9% (채택)",options:{color:PAPER}}],{x:MX+bw+0.58,y:yy+0.5,w:bw-0.36,h:0.75,margin:0,fontFace:BODY,fontSize:9.5,valign:"top",lineSpacingMultiple:1.1});
    s.addText("EV 효과: ex-AMPC EBITDA 5%→9%로 DCF EV $1.83bn → $3.45bn(NOL 포함).",{x:MX,y:H-0.62,w:8,h:0.2,margin:0,fontFace:BODY,fontSize:8,color:RED,valign:"middle"});
    footer(s,"10");
  }

  await pres.writeFile({ fileName: require("path").join(__dirname,"..","output","SK_Valuation_v2.pptx") });
  console.log("WROTE v2 deck");
})();
