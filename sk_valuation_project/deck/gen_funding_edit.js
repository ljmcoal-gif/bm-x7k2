// SK Funding Deck — 편집 가능(native) 버전
// 이미지가 아니라 pptxgenjs 네이티브 텍스트/표/도형으로 5슬라이드 생성 → 파워포인트에서 직접 편집 가능.
// 디자인: SK 실적발표 템플릿(명조 헤드라인 + IBM Plex + Space Grotesk 수치).
// 실행: NODE_PATH=$(npm root -g) node gen_funding_edit.js  → output/SK_Funding_v3_editable.pptx
const groot = require('child_process').execSync('npm root -g').toString().trim();
const pptxgen = require(groot + '/pptxgenjs');
const path = require('path');
const pres = new pptxgen();

const W = 11.69, H = 8.27;
pres.defineLayout({ name: 'A4L', width: W, height: H });
pres.layout = 'A4L';

// ── 토큰 ──
const BG='FBF9F5', INK='1B1714', DARK='211C18', RED='C1121F', WARM='D8CFC2', GREEN='1E7A4D',
      GRID='ECE6DC', RULE='E4DED4', MUTE='6F6960', MUTE2='9A847C', PAPER='FFFFFF', REDSOFT='FBF1EF',
      GREENSOFT='E8F3EC', SAND='F2ECE0', REDLT='FF8E84', GREENLT='86D6A6';
const HEAD='Noto Serif KR', BODY='IBM Plex Sans KR', MONO='IBM Plex Mono', SG='Space Grotesk';
const MX=0.62, TOP=0.52;
const RR = pres.shapes.ROUNDED_RECTANGLE, RECT = pres.shapes.RECTANGLE, LINE = pres.shapes.LINE, OVAL = pres.shapes.OVAL;

function base(s){ s.background = { color: BG }; }
function header(s, eyebrow, page, title, dek){
  s.addText(eyebrow, { x:MX, y:TOP, w:7.5, h:0.26, margin:0, fontFace:MONO, fontSize:9, color:RED, charSpacing:2, valign:'middle' });
  s.addText('SK Innovation · IR  /  '+page, { x:W-4.5-MX, y:TOP, w:4.5, h:0.26, margin:0, align:'right', fontFace:MONO, fontSize:9, color:MUTE2, charSpacing:1.5, valign:'middle' });
  s.addText(title, { x:MX, y:TOP+0.34, w:W-2*MX-2.6, h:0.62, margin:0, fontFace:HEAD, fontSize:26, bold:true, color:INK, valign:'middle' });
  if(dek) s.addText(dek, { x:W-MX-3.0, y:TOP+0.5, w:3.0, h:0.42, margin:0, align:'right', fontFace:BODY, fontSize:9.5, color:MUTE, valign:'bottom' });
  s.addShape(LINE, { x:MX, y:TOP+1.04, w:W-2*MX, h:0, line:{ color:INK, width:1.5 } });
}
function footer(s){
  s.addShape(LINE, { x:MX, y:H-0.46, w:W-2*MX, h:0, line:{ color:RULE, width:0.75 } });
  s.addText('SK 전략팀 · Strictly Private & Confidential', { x:MX, y:H-0.42, w:6, h:0.26, margin:0, fontFace:BODY, fontSize:7.5, color:MUTE2, valign:'middle' });
  s.addText('★ = 미검증', { x:W-MX-2.5, y:H-0.42, w:2.5, h:0.26, margin:0, align:'right', fontFace:MONO, fontSize:7.5, color:MUTE2, valign:'middle' });
}
function eb(s,x,y,w,t){ s.addText(t,{x,y,w,h:0.24,margin:0,fontFace:MONO,fontSize:8.5,color:MUTE2,charSpacing:1.5,valign:'middle'}); }

(async()=>{

// ════════ S1 · 표지 ════════
{
  const s=pres.addSlide(); base(s);
  s.addText('F',{x:W-4.6,y:0.6,w:4.4,h:5.0,margin:0,align:'right',fontFace:HEAD,fontSize:340,bold:true,color:RED,transparency:94,valign:'top'});
  s.addShape(RECT,{x:MX,y:0.62,w:0.42,h:0.06,fill:{color:RED}});
  s.addText('SK 전략팀',{x:MX+0.5,y:0.5,w:4,h:0.3,margin:0,fontFace:BODY,bold:true,fontSize:14,color:INK,valign:'middle'});
  s.addText('FINANCING STRUCTURE',{x:W-MX-4,y:0.5,w:4,h:0.3,margin:0,align:'right',fontFace:MONO,fontSize:10,color:MUTE2,charSpacing:3,valign:'middle'});
  s.addText('FUNDING STRUCTURE v2',{x:MX,y:H-3.55,w:8,h:0.34,margin:0,fontFace:MONO,fontSize:13,color:RED,charSpacing:4,valign:'middle'});
  s.addText('펀딩 구조\n전면 개편(안)',{x:MX-0.02,y:H-3.05,w:9,h:1.85,margin:0,fontFace:HEAD,fontSize:54,bold:true,color:INK,lineSpacingMultiple:1.02,valign:'top'});
  s.addText('BA 보조금-우선 · OT 링펜스 유지 — 재무팀 의견 반영',{x:MX,y:H-1.2,w:9,h:0.3,margin:0,fontFace:BODY,fontSize:14,color:MUTE,valign:'middle'});
  s.addShape(LINE,{x:MX,y:H-0.82,w:W-2*MX,h:0,line:{color:RULE,width:1}});
  s.addText('2026년 6월',{x:MX,y:H-0.75,w:4,h:0.3,margin:0,fontFace:BODY,fontSize:11,color:MUTE,valign:'middle'});
  s.addText('SK 전략팀 · Internal',{x:W-MX-4,y:H-0.75,w:4,h:0.3,margin:0,align:'right',fontFace:MONO,fontSize:11,color:MUTE,valign:'middle'});
}

// ════════ S2 · 현황 & 개조 ════════
{
  const s=pres.addSlide(); base(s);
  header(s,'CURRENT STATE & RETROFIT','02','현황 & 개조 소요','BA · OT 재무상태 → 개조');
  const colL=MX, colLW=5.55, colR=MX+colLW+0.4, colRW=W-colR-MX;
  eb(s,colL,1.78,colLW,'주요 재무상태 · KEY FINANCIALS');
  // 재무 표 (editable addTable)
  const th={fontFace:BODY,fontSize:11,bold:true,color:PAPER,valign:'middle'};
  const kk={fontFace:BODY,fontSize:9.5,color:MUTE,valign:'middle'};
  const vv={fontFace:MONO,fontSize:9.5,bold:true,color:INK,align:'right',valign:'middle'};
  function finTbl(y,title,tint,sub,rows){
    s.addTable([[{text:title,options:{...th,fill:{color:tint}}},{text:sub,options:{fontFace:MONO,fontSize:8,color:WARM,align:'right',valign:'middle',fill:{color:tint}}}]],
      {x:colL,y,w:colLW,colW:[3.5,colLW-3.5],rowH:0.42,margin:[3,8,3,8]});
    const body=rows.map(r=>[{text:r[0],options:kk},{text:r[1],options:{...vv,color:r[2]||INK}}]);
    s.addTable(body,{x:colL,y:y+0.42,w:colLW,colW:[3.0,colLW-3.0],rowH:0.365,border:{type:'solid',color:GRID,pt:0.5},margin:[2,8,2,8],fill:{color:PAPER}});
  }
  finTbl(2.06,'공장 1 — SKBA (조지아)',INK,'파우치 · 모회사',
    [['명판 / 가동(COD)','21.2 GWh / 2027~28'],['기존 차입 (단독)','~$5.0bn + SKI 보증',RED],['run-rate 매출','~$2.2bn'],['ex-AMPC EBITDA','9% (목표)'],['순자산 NAV ★','~$2.0bn']]);
  finTbl(4.28,'공장 2 — SKOT (테네시)',RED,'BA 완전자회사 100%',
    [['명판 / 가동(COD)','29.7 GWh / 2028'],['기존 차입 (DOE)','$4.0bn ATVM (이자만)',RED],['run-rate 매출','~$2.7bn'],['ex-AMPC EBITDA','9% (목표)'],['순자산 NAV ★','~$3.0bn']]);
  // 우: 개조 금액 (다크 박스)
  eb(s,colR,1.78,colRW,'개조 금액 · RETROFIT CAPEX');
  s.addShape(RR,{x:colR,y:2.06,w:colRW,h:1.5,rectRadius:0.06,fill:{color:DARK},line:{type:'none'}});
  const cap=[['BA (개조)','$0.4bn',PAPER],['OT','$1.0bn',PAPER]];
  let cy=2.24;
  cap.forEach(c=>{ s.addText(c[0],{x:colR+0.3,y:cy,w:2,h:0.34,margin:0,fontFace:BODY,fontSize:11,color:WARM,valign:'middle'});
    s.addText(c[1],{x:colR+colRW-2.3,y:cy,w:2,h:0.34,margin:0,align:'right',fontFace:MONO,fontSize:20,bold:true,color:c[2],valign:'middle'}); cy+=0.42; });
  s.addShape(LINE,{x:colR+0.3,y:cy+0.02,w:colRW-0.6,h:0,line:{color:'4a443d',width:0.75}});
  s.addText('합계 소요',{x:colR+0.3,y:cy+0.1,w:2,h:0.38,margin:0,fontFace:BODY,bold:true,fontSize:12,color:PAPER,valign:'middle'});
  s.addText('$1.4bn',{x:colR+colRW-2.5,y:cy+0.1,w:2.2,h:0.38,margin:0,align:'right',fontFace:MONO,fontSize:24,bold:true,color:REDLT,valign:'middle'});
  // 기대효과
  eb(s,colR,3.78,colRW,'기대 효과 · EXPECTED EFFECT');
  const eff=['DC블록 ASP $155/kWh 사업화 — 셀+모듈+컨테이너','ex-AMPC EBITDA 9% → 지분 수익률(FCFE) 충족','DCF EV ≈ $3.6bn (밴드 $3.0~4.0bn)','NAV(SOTP) ≈ $5.0bn ★'];
  let ey=4.08;
  eff.forEach(t=>{ s.addText('●',{x:colR,y:ey,w:0.25,h:0.3,margin:0,fontFace:BODY,fontSize:11,color:RED,valign:'top'});
    s.addText(t,{x:colR+0.28,y:ey,w:colRW-0.28,h:0.4,margin:0,fontFace:BODY,fontSize:10.5,color:INK,valign:'top'}); ey+=0.44; });
  // KPI 2x2 (Space Grotesk)
  const kp=[['총 명판','50.9','GWh',INK],['기존 차입','9.0','bn',INK],['개조 소요','1.4','bn',RED],['ASP','155','$/kWh',INK]];
  const kw=(colRW-0.4)/2; let kx=colR, ky=6.0;
  kp.forEach((k,i)=>{ const x=colR+(i%2)*(kw+0.4), y=6.0+Math.floor(i/2)*1.02;
    s.addShape(LINE,{x,y,w:kw,h:0,line:{color:k[3]===RED?RED:INK,width:2}});
    s.addText(k[0],{x,y:y+0.06,w:kw,h:0.24,margin:0,fontFace:BODY,fontSize:10,color:MUTE,valign:'middle'});
    s.addText([{text:k[1],options:{fontSize:26}},{text:k[2],options:{fontSize:12}}],{x,y:y+0.32,w:kw,h:0.5,margin:0,fontFace:SG,bold:true,color:k[3],valign:'middle'}); });
  footer(s);
}

// ════════ S3 · 펀딩 제안 (지분구조 플로우) ════════
{
  const s=pres.addSlide(); base(s);
  header(s,'FUNDING PROPOSAL','03','펀딩 제안 — BA 중심','지분구조 · 캐시 플로우');
  const colL=MX, colLW=6.9, colR=MX+colLW+0.35, colRW=W-colR-MX;
  eb(s,colL,1.78,colLW,'지분구조 & 캐시 플로우 · STRUCTURE & CASH');
  // 엔티티 체인 (우측 정렬 spine)
  const boxW=2.5, boxX=colL+colLW-boxW-0.3;
  function entBox(y,fill,nm,sub,txtc){ s.addShape(RR,{x:boxX,y,w:boxW,h:0.62,rectRadius:0.05,fill:{color:fill},line:{type:'none'}});
    s.addText(nm,{x:boxX,y:y+0.06,w:boxW,h:0.3,margin:0,align:'center',fontFace:BODY,bold:true,fontSize:13,color:txtc||PAPER,valign:'middle'});
    s.addText(sub,{x:boxX,y:y+0.34,w:boxW,h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:8,color:WARM,valign:'middle'}); }
  function arrow(y,txt,c){ s.addText('▼',{x:boxX+boxW/2-0.3,y,w:0.6,h:0.24,margin:0,align:'center',fontFace:BODY,fontSize:13,color:c||INK,valign:'middle'});
    s.addText(txt,{x:boxX+boxW/2+0.15,y,w:2,h:0.24,margin:0,fontFace:MONO,fontSize:8.5,color:c||MUTE,bold:!!c,valign:'middle'}); }
  function pill(y,txt,val,tint,brd,txtc){ const pw=3.0, px=boxX-pw-0.35;
    s.addShape(RR,{x:px,y,w:pw,h:0.44,rectRadius:0.22,fill:{color:tint||PAPER},line:{color:brd||RULE,width:1.5}});
    s.addText([{text:txt+'  ',options:{color:txtc||MUTE}},{text:val,options:{fontFace:MONO,bold:true,color:txtc||INK}}],{x:px+0.18,y,w:pw-0.5,h:0.44,margin:0,fontFace:BODY,fontSize:9.5,valign:'middle'});
    s.addText('▶',{x:px+pw-0.02,y,w:0.4,h:0.44,margin:0,align:'center',fontFace:BODY,fontSize:11,color:RED,valign:'middle'}); }
  entBox(2.06,'4A443D','SK이노베이션','최상위 모회사'); arrow(2.74,'100%');
  entBox(3.02,'5E5346','SK온 (SK On)','배터리 중간지주'); arrow(3.70,'100%');
  // BA + cash pills
  entBox(4.10,INK,'BA — SKBA · 조달 주체','총 $1.4bn 일괄 조달');
  pill(3.90,'보조금 (10%)','$0.14bn',GREENSOFT,GREEN,GREEN);
  pill(4.42,'비소구부채 (60%)','$0.85bn');
  s.addText('▼',{x:boxX+boxW/2-0.3,y:4.86,w:0.6,h:0.24,margin:0,align:'center',fontFace:BODY,fontSize:13,color:RED,valign:'middle'});
  s.addText('$1.0bn 증자 (equity-down)',{x:boxX-2.4,y:4.86,w:2.35,h:0.24,margin:0,align:'right',fontFace:MONO,fontSize:8.5,color:RED,bold:true,valign:'middle'});
  entBox(5.22,RED,'OT — SKOT','완전자회사 100%');
  pill(5.30,'기존 DOE ATVM','$4.0bn');
  s.addText('조달 믹스 그랜트 10% · 지분 30% · 비소구부채 60% ($1.4bn) → 자체 $0.4bn + OT $1.0bn 증자',
    {x:colL,y:6.06,w:colLW,h:0.3,margin:0,align:'center',fontFace:BODY,fontSize:9.5,color:MUTE,valign:'middle'});
  // 우: 제안
  eb(s,colR,1.78,colRW,'제안 · PROPOSAL');
  function prop(y,pn,pt,ps){ s.addShape(RR,{x:colR,y,w:colRW,h:1.85,rectRadius:0.06,fill:{color:REDSOFT},line:{color:RED,width:1.5}});
    s.addText(pn,{x:colR+0.26,y:y+0.16,w:colRW-0.5,h:0.26,margin:0,fontFace:MONO,fontSize:10,color:MUTE2,charSpacing:1.5,valign:'middle'});
    s.addText(pt,{x:colR+0.26,y:y+0.42,w:colRW-0.5,h:0.4,margin:0,fontFace:HEAD,bold:true,fontSize:17,color:RED,valign:'middle'});
    s.addText(ps,{x:colR+0.26,y:y+0.86,w:colRW-0.5,h:0.9,margin:0,fontFace:BODY,fontSize:10,color:MUTE,valign:'top',lineSpacingMultiple:1.15}); }
  prop(2.06,'제안 ①','BA가 $1.4bn 일괄 조달','OT는 DOE·CoC로 신규 차입이 어렵다 → BA가 전액 조달해 OT엔 $1.0bn 증자(equity-down). 자회사 자금주입의 가장 깨끗한 경로 — 신규 OT 부채·DSCR 불요.');
  prop(4.06,'제안 ②','그랜트 극대화 요청','BA 개조비 $0.4bn을 정부 보조금(MESC·48C)으로 최대한. 무상환·무희석·무CoC → EBIT=0 자산에 깨끗한 자금.');
  // 결론 바
  s.addShape(RR,{x:MX,y:6.5,w:W-2*MX,h:0.72,rectRadius:0.05,fill:{color:SAND},line:{type:'none'}});
  s.addText([{text:'[결론]  ',options:{bold:true,color:RED}},{text:'BA가 $1.4bn 일괄 조달 → 자체 $0.4bn + OT 증자 $1.0bn. 조달은 ',options:{color:INK}},{text:'그랜트 최우선',options:{bold:true,color:RED}},{text:', 부족분만 비소구부채·Equity. ($1.4bn이 BA에 집중 — 조달여력 확인 ★)',options:{color:INK}}],
    {x:MX+0.25,y:6.5,w:W-2*MX-0.5,h:0.72,margin:0,fontFace:BODY,fontSize:10.5,valign:'middle',lineSpacingMultiple:1.1});
  footer(s);
}

// ════════ S4 · 지분 성립 조건 (역산) ════════
{
  const s=pres.addSlide(); base(s);
  header(s,'VALUATION — REQUIRED THRESHOLD','04','지분 성립 조건 — 역산','20% 희석 한도 기준');
  const colL=MX, colLW=4.2, colR=MX+colLW+0.5, colRW=W-colR-MX;
  eb(s,colL,1.78,colLW,'주요 가정 · KEY ASSUMPTIONS');
  const asm=[['DC블록 ASP','$155 /kWh'],['공급 셀 (개조 후)','파우치 + 각형'],['신규 조달','$1.4bn (10/30/60)'],['외부 지분 (30%)','$0.43bn'],['순차입금 (BA/OT)','$6.4bn (BA3.7+OT2.6)'],['EBITDA ex-AMPC (9%)','$441mm run-rate'],['peer EV/EBITDA','8~13x']];
  let ay=2.1;
  asm.forEach(a=>{ s.addText(a[0],{x:colL,y:ay,w:2.2,h:0.42,margin:0,fontFace:BODY,fontSize:10,color:MUTE,valign:'middle'});
    s.addText(a[1],{x:colL+2.1,y:ay,w:colLW-2.1,h:0.42,margin:0,align:'right',fontFace:MONO,fontSize:10.5,bold:true,color:INK,valign:'middle'});
    s.addShape(LINE,{x:colL,y:ay+0.44,w:colLW,h:0,line:{color:RULE,width:0.5}}); ay+=0.475; });
  // 우: 역산 체인
  eb(s,colR,1.78,colRW,'역산 — 20% 희석 한도를 지키려면 · THE BAR');
  const steps=[['1','최소 지분가치','외부 지분 $0.43bn ÷ 20% 희석','$2.1','bn',INK,false],
    ['2','→ 최소 EV','지분 $2.1 + 순차입금 $6.4 (BA3.7+OT2.6)','$8.6','bn',INK,false],
    ['3','→ 최소 implied 멀티플','현 EBITDA $441mm · peer 8~13x 초과','~19','x',RED,false],
    ['4','→ DCF 환산: 필요 ex-AMPC EBITDA','현 9%의 약 2.5배 (모델 sweep)','~22','%',RED,true]];
  let sy=2.1;
  steps.forEach(st=>{ const hl=st[6];
    s.addShape(RR,{x:colR,y:sy,w:colRW,h:0.82,rectRadius:0.06,fill:{color:hl?REDSOFT:PAPER},line:{color:hl?RED:RULE,width:hl?1.5:1}});
    s.addShape(OVAL,{x:colR+0.2,y:sy+0.24,w:0.34,h:0.34,fill:{color:hl?RED:WARM},line:{type:'none'}});
    s.addText(st[0],{x:colR+0.2,y:sy+0.24,w:0.34,h:0.34,margin:0,align:'center',fontFace:HEAD,bold:true,fontSize:14,color:hl?PAPER:INK,valign:'middle'});
    s.addText(st[1],{x:colR+0.68,y:sy+0.12,w:colRW-2.5,h:0.34,margin:0,fontFace:BODY,bold:true,fontSize:12.5,color:hl?RED:INK,valign:'middle'});
    s.addText(st[2],{x:colR+0.68,y:sy+0.45,w:colRW-2.5,h:0.28,margin:0,fontFace:BODY,fontSize:9,color:MUTE,valign:'middle'});
    s.addText([{text:st[3],options:{fontSize:26}},{text:st[4],options:{fontSize:14}}],{x:colR+colRW-1.9,y:sy+0.14,w:1.7,h:0.55,margin:0,align:'right',fontFace:SG,bold:true,color:st[5],valign:'middle'}); sy+=0.92; });
  // 게이지
  s.addShape(RR,{x:MX,y:5.95,w:W-2*MX,h:0.72,rectRadius:0.05,fill:{color:DARK},line:{type:'none'}});
  s.addText('필요 마진 갭 — 9% → ~22%',{x:MX+0.3,y:5.95,w:2.7,h:0.72,margin:0,fontFace:HEAD,bold:true,fontSize:12,color:REDLT,valign:'middle'});
  const gx=MX+3.3, gw=W-MX-gx-0.3, gy=6.31;
  s.addShape(LINE,{x:gx,y:gy,w:gw,h:0,line:{color:'3a342e',width:3}});
  s.addShape(LINE,{x:gx+gw*0.36,y:gy,w:gw*0.54,h:0,line:{color:RED,width:3}});
  s.addText('현 9%',{x:gx+gw*0.36-0.5,y:gy-0.32,w:1,h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:9,color:WARM,valign:'middle'});
  s.addText('필요 ~22%',{x:gx+gw*0.90-0.7,y:gy-0.32,w:1.4,h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:9,bold:true,color:REDLT,valign:'middle'});
  // 결론 콜아웃
  s.addShape(RR,{x:MX,y:6.82,w:W-2*MX,h:0.86,rectRadius:0.06,fill:{color:DARK},line:{type:'none'}});
  s.addText('결론 — 마진이 딜의 성패를 가른다',{x:MX+0.3,y:6.92,w:6,h:0.28,margin:0,fontFace:HEAD,bold:true,fontSize:13,color:REDLT,valign:'middle'});
  s.addText([{text:'현 9% 마진으로는 20% 희석 한도 내 지분조달 불가. 길은 둘 — 시장이 ',options:{color:WARM}},{text:'~19x(peer 상단 초과)',options:{color:REDLT,bold:true}},{text:'를 주거나, ',options:{color:WARM}},{text:'ex-AMPC EBITDA를 ~22%로',options:{color:GREENLT,bold:true}},{text:' 올려 정상 멀티플(≈8x)에서 $8.6bn EV를 정당화하거나. ★ 순차입금 $6.4bn · 모델 sweep',options:{color:WARM}}],
    {x:MX+0.3,y:7.2,w:W-2*MX-0.6,h:0.42,margin:0,fontFace:BODY,fontSize:9.5,valign:'top',lineSpacingMultiple:1.1});
  footer(s);
}

// ════════ S5 · 커뮤니케이션 (마무리) ════════
{
  const s=pres.addSlide(); base(s);
  header(s,'CLOSING — GOVERNMENT ENGAGEMENT','05','커뮤니케이션 방식','재무 사전 정렬 우선');
  eb(s,MX,1.78,7,'진행 순서 · ENGAGEMENT SEQUENCE');
  const tw=(W-2*MX-3*0.3)/4;
  const tsteps=[['1','내부 재무 사전 정렬','ask·구조·레드라인 재무 합의 (선행)',true],['2','ask 확정 → 매칭','요청 확정 후 카운터파트로',false],['3','DOE 부서 미팅','합의된 ask · 부서별 단일 메시지',false],['4','DOW(국방) 트랙','방산·딥테크 별도 진행',false]];
  tsteps.forEach((t,i)=>{ const x=MX+i*(tw+0.3), y=2.08, hl=t[3];
    s.addShape(RR,{x,y,w:tw,h:1.15,rectRadius:0.06,fill:{color:hl?DARK:PAPER},line:{color:hl?DARK:RULE,width:1.5}});
    s.addShape(OVAL,{x:x+0.24,y:y+0.22,w:0.42,h:0.42,fill:{color:hl?RED:WARM},line:{type:'none'}});
    s.addText(t[0],{x:x+0.24,y:y+0.22,w:0.42,h:0.42,margin:0,align:'center',fontFace:HEAD,bold:true,fontSize:15,color:hl?PAPER:INK,valign:'middle'});
    s.addText(t[1],{x:x+0.24,y:y+0.72,w:tw-0.48,h:0.24,margin:0,fontFace:BODY,bold:true,fontSize:11,color:hl?PAPER:INK,valign:'middle'});
    s.addText(t[2],{x:x+0.24,y:y+0.95,w:tw-0.48,h:0.16,margin:0,fontFace:BODY,fontSize:7.5,color:hl?WARM:MUTE,valign:'middle'});
    if(i<3) s.addText('▶',{x:x+tw-0.02,y:y+0.4,w:0.32,h:0.3,margin:0,align:'center',fontFace:BODY,fontSize:13,color:RED,valign:'middle'}); });
  // 하: 부서매칭 표 + 레드라인
  const colL=MX, colLW=5.4, colR=MX+colLW+0.5, colRW=W-colR-MX;
  eb(s,colL,3.6,colLW,'ASK → DOE 카운터파트 매칭');
  const mk={fontFace:BODY,fontSize:10,bold:true,color:INK,valign:'middle'};
  const mrows=[[{text:'보조금 (BA)',options:mk},{text:'MESC — 제조·공급망',options:{fontFace:BODY,fontSize:9.5,color:GREEN,valign:'middle'}}],
    [{text:'대출',options:mk},{text:'LPO (ATVM/Title 17) · OT 기존',options:{fontFace:BODY,fontSize:9.5,color:INK,valign:'middle'}}],
    [{text:'Equity',options:mk},{text:'해당 부서 없음 — DOE 지분투자 ✗',options:{fontFace:BODY,fontSize:9.5,color:RED,valign:'middle'}}],
    [{text:'방산 앵글',options:mk},{text:'DOW (국방) 별도 트랙',options:{fontFace:BODY,fontSize:9.5,color:INK,valign:'middle'}}]];
  s.addTable(mrows,{x:colL,y:3.9,w:colLW,colW:[1.7,colLW-1.7],rowH:0.52,border:{type:'solid',color:RULE,pt:0.5},margin:[3,6,3,6]});
  // 레드라인 (다크)
  eb(s,colR,3.6,colRW,'레드라인 · RED LINES');
  s.addShape(RR,{x:colR,y:3.9,w:colRW,h:2.08,rectRadius:0.06,fill:{color:DARK},line:{type:'none'}});
  s.addText('미팅 전 재무 합의 사항',{x:colR+0.26,y:4.04,w:colRW-0.5,h:0.32,margin:0,fontFace:HEAD,bold:true,fontSize:13,color:PAPER,valign:'middle'});
  s.addText([{text:'• SK이노 보증 확대 금지\n',options:{color:WARM}},{text:'• BA 기존 covenant breach 금지\n',options:{color:WARM}},{text:'• CoC 20% 캡 유지\n',options:{color:WARM}},{text:'• DOE Equity 기대 제외 (대출·보증만)',options:{color:REDLT,bold:true}}],
    {x:colR+0.26,y:4.42,w:colRW-0.5,h:1.5,margin:0,fontFace:BODY,fontSize:11,valign:'top',lineSpacingMultiple:1.4});
  // 예상 타임라인 strip
  const mil=[['D+0 · 1~2주','재무 사전 정렬 메모',RED],['D+2~3주','ask 확정 · 부서 매칭',MUTE2],['D+4~6주','DOE 부서 미팅',MUTE2],['D+8주~','DOW·후속 협상',MUTE2]];
  mil.forEach((m,i)=>{ const x=MX+i*(tw+0.3), y=6.3;
    s.addShape(LINE,{x,y,w:tw,h:0,line:{color:INK,width:2}});
    s.addText(m[0],{x,y:y+0.08,w:tw,h:0.22,margin:0,fontFace:MONO,fontSize:9,bold:true,color:m[2],valign:'middle'});
    s.addText(m[1],{x,y:y+0.3,w:tw,h:0.24,margin:0,fontFace:BODY,bold:true,fontSize:10,color:INK,valign:'middle'}); });
  s.addShape(RR,{x:MX,y:7.05,w:W-2*MX,h:0.5,rectRadius:0.04,fill:{color:SAND},line:{type:'none'}});
  s.addText([{text:'[원칙]  ',options:{bold:true,color:RED}},{text:"부서 미팅 전, '무엇을 요청할지'를 재무와 먼저 확정한다 — 잘못된 부서에 잘못된 ask로 가는 것을 막는다.",options:{color:INK}}],
    {x:MX+0.25,y:7.05,w:W-2*MX-0.5,h:0.5,margin:0,fontFace:BODY,fontSize:9.5,valign:'middle'});
  footer(s);
}

  await pres.writeFile({ fileName: path.join(__dirname,'..','output','SK_Funding_v3_editable.pptx') });
  console.log('WROTE editable pptx');
})();
