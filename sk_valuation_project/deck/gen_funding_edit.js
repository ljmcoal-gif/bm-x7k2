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
    [['명판 / 가동(COD)','21.2 GWh / 2027~28'],['기존 차입 (단독)','~$5.0bn + SKI 보증',RED],['run-rate 매출','~$2.2bn'],['개조 ex-AMPC EBITDA','9% (목표)'],['기저 EBITDA (pre-SOP) ★','$0.2bn (2026~)',GREEN],['순자산 NAV ★','~$2.0bn']]);
  finTbl(4.82,'공장 2 — SKOT (테네시)',RED,'BA 완전자회사 100%',
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
  const eff=['DC블록 ASP $155/kWh 사업화 — 셀+모듈+컨테이너','SKBA pre-SOP 기저 $0.2bn(2026~) · run-rate ex-AMPC $442mm 유지','DCF EV ≈ $4.0bn (기저 조기현금 반영·밴드 $3.5~4.5bn)','NAV(SOTP) ≈ $5.0bn ★'];
  let ey=4.08;
  eff.forEach(t=>{ s.addText('●',{x:colR,y:ey,w:0.25,h:0.3,margin:0,fontFace:BODY,fontSize:11,color:RED,valign:'top'});
    s.addText(t,{x:colR+0.28,y:ey,w:colRW-0.28,h:0.4,margin:0,fontFace:BODY,fontSize:10.5,color:INK,valign:'top'}); ey+=0.44; });
  // KPI 2x2 (Space Grotesk)
  const kp=[['총 명판','50.9','GWh',INK],['run-rate EBITDA','442','$mm',GREEN],['개조 소요','1.4','bn',RED],['DCF EV','4.0','bn',RED]];
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

// ════════ S4 · 지분 성립 조건 (조달 시나리오) ════════
{
  const s=pres.addSlide(); base(s);
  header(s,'VALUATION — FUNDING SCENARIOS','04','지분 성립 조건 — 조달 시나리오','희석·그랜트 레버 → implied 멀티플');
  const colL=MX, colLW=3.5, colR=MX+colLW+0.4, colRW=W-colR-MX;
  // 좌: 가정 + 역산식
  eb(s,colL,1.78,colLW,'주요 가정 · KEY ASSUMPTIONS');
  const asm=[['DC블록 ASP','$155/kWh'],['공급 셀 (개조 후)','파우치+각형'],['SKBA 기저 EBITDA ★','$0.2bn pre-SOP',GREEN],['run-rate ex-AMPC','$442mm',INK],['신규 조달','$1.42bn'],['순차입금 (BA/OT)','$6.4bn'],['peer EV/EBITDA','8~13x']];
  let ay=2.08;
  asm.forEach(a=>{ s.addText(a[0],{x:colL,y:ay,w:2.0,h:0.36,margin:0,fontFace:BODY,fontSize:9.5,color:MUTE,valign:'middle'});
    s.addText(a[1],{x:colL+1.7,y:ay,w:colLW-1.7,h:0.36,margin:0,align:'right',fontFace:MONO,fontSize:10,bold:true,color:a[2]||INK,valign:'middle'});
    s.addShape(LINE,{x:colL,y:ay+0.37,w:colLW,h:0,line:{color:RULE,width:0.5}}); ay+=0.40; });
  // 역산식 박스
  s.addShape(RR,{x:colL,y:ay+0.06,w:colLW,h:1.28,rectRadius:0.06,fill:{color:DARK},line:{type:'none'}});
  s.addText('역산 논리 · THE MECHANISM',{x:colL+0.22,y:ay+0.16,w:colLW-0.4,h:0.24,margin:0,fontFace:MONO,fontSize:8,color:WARM,charSpacing:1,valign:'middle'});
  s.addText([{text:'최소지분',options:{color:REDLT,bold:true}},{text:' = 외부지분 ÷ 희석캡\n',options:{color:WARM}},{text:'최소EV',options:{color:REDLT,bold:true}},{text:' = 최소지분 + 순차입금\n',options:{color:WARM}},{text:'멀티플',options:{color:REDLT,bold:true}},{text:' = 최소EV ÷ run-rate EBITDA',options:{color:WARM}}],
    {x:colL+0.22,y:ay+0.42,w:colLW-0.4,h:0.86,margin:0,fontFace:BODY,fontSize:9.5,valign:'top',lineSpacingMultiple:1.25});
  // 우: 시나리오 표 (editable)
  eb(s,colR,1.78,colRW,'조달 시나리오 — implied min EV/EBITDA · SCENARIOS');
  const hd={fontFace:BODY,fontSize:9,bold:true,color:PAPER,align:'center',valign:'middle',fill:{color:INK}};
  const hdl={...hd,align:'left'};
  const head4=[{text:'시나리오',options:hdl},{text:'그랜트',options:hd},{text:'희석캡',options:hd},{text:'최소 EV',options:hd},{text:'멀티플',options:hd},{text:'peer초과',options:hd}];
  const cL={fontFace:BODY,fontSize:9.5,color:INK,valign:'middle'};
  const cC={fontFace:MONO,fontSize:9.5,color:INK,align:'center',valign:'middle'};
  const cM={fontFace:SG,fontSize:12,bold:true,align:'center',valign:'middle'};
  function srow(nm,g,d,ev,mult,tag,mc,tc,fill){
    return [{text:nm,options:{...cL,fill:{color:fill}}},{text:g,options:{...cC,fill:{color:fill}}},{text:d,options:{...cC,fill:{color:fill}}},
      {text:ev,options:{...cC,fill:{color:fill}}},{text:mult,options:{...cM,color:mc,fill:{color:fill}}},{text:tag,options:{fontFace:MONO,fontSize:8,bold:true,color:tc,align:'center',valign:'middle',fill:{color:fill}}}];
  }
  const body4=[head4,
    srow('BASE  (그랜트10/지분30)','10%','20%','$8.6bn','19.3x','+6.3x',RED,RED,PAPER),
    srow('① CoC 20% + 그랜트 20%','20%','20%','$7.8bn','17.7x','+4.7x',RED,MUTE,GREENSOFT),
    srow('② CoC 해제 → 희석 30%','10%','30%','$7.8bn','17.7x','+4.7x',RED,MUTE,GREENSOFT),
    srow('①+② 결합','20%','30%','$7.3bn','16.6x','+3.6x',RED,MUTE,GREENSOFT)];
  s.addTable(body4,{x:colR,y:2.08,w:colRW,colW:[2.35,0.72,0.72,0.95,0.85,0.86],rowH:[0.34,0.56,0.56,0.56,0.56],border:{type:'solid',color:RULE,pt:0.5},valign:'middle',margin:[2,5,2,5]});
  // peer 밴드 바
  eb(s,colR,4.9,colRW,'peer 밴드 대비 위치 · 조달 레버는 갭을 좁힌다');
  const bx=colR+0.05, bw=colRW-0.7, by=5.55, lo=8, hi=20;
  const px=v=>bx+bw*(v-lo)/(hi-lo);
  s.addShape(LINE,{x:bx,y:by,w:bw,h:0,line:{color:WARM,width:3}});
  // peer band 8~13 (green tint segment)
  s.addShape(RECT,{x:px(8),y:by-0.09,w:px(13)-px(8),h:0.18,fill:{color:GREENSOFT},line:{color:GREEN,width:0.75}});
  s.addText('peer 8~13x',{x:px(8),y:by-0.34,w:px(13)-px(8),h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:8,color:GREEN,valign:'middle'});
  // 개선 방향 화살표 (BASE 19.3 → ①+② 16.6)
  s.addShape(LINE,{x:px(16.6),y:by-0.24,w:px(19.3)-px(16.6),h:0,line:{color:REDLT,width:1.25,endArrowType:'triangle',beginArrowType:'none'}});
  s.addText('조달 레버 −2.7x',{x:px(16.6)-0.3,y:by-0.44,w:px(19.3)-px(16.6)+0.6,h:0.18,margin:0,align:'center',fontFace:MONO,fontSize:7,color:REDLT,valign:'middle'});
  [[19.3,'BASE',RED],[17.7,'①·②','8A5A55'],[16.6,'①+②','8A5A55']].forEach(m=>{
    s.addShape(OVAL,{x:px(m[0])-0.05,y:by-0.05,w:0.1,h:0.1,fill:{color:m[2]},line:{type:'none'}});
    s.addText(m[1],{x:px(m[0])-0.5,y:by+0.1,w:1.0,h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:7.5,bold:true,color:m[2],valign:'middle'});
    s.addText(m[0]+'x',{x:px(m[0])-0.5,y:by-0.02,w:1.0,h:0.2,margin:0,align:'center',fontFace:MONO,fontSize:8,bold:true,color:m[2],valign:'top'}); });
  [8,13,16,20].forEach(t=>s.addText(t+'x',{x:px(t)-0.3,y:by+0.28,w:0.6,h:0.18,margin:0,align:'center',fontFace:MONO,fontSize:7,color:MUTE2,valign:'middle'}));
  // 결론 콜아웃 (다크)
  s.addShape(RR,{x:MX,y:6.34,w:W-2*MX,h:1.28,rectRadius:0.06,fill:{color:DARK},line:{type:'none'}});
  s.addText('결론 — 조달 레버는 갭을 좁히지만, 밴드 진입은 마진이 결정한다',{x:MX+0.3,y:6.44,w:W-2*MX-0.6,h:0.3,margin:0,fontFace:HEAD,bold:true,fontSize:13,color:REDLT,valign:'middle'});
  s.addText([{text:'run-rate ex-AMPC EBITDA ',options:{color:WARM}},{text:'$442mm 유지',options:{color:GREENLT,bold:true}},{text:' 기준(기저 $0.2bn은 pre-SOP 조기현금·EV 방어용, run-rate 불변). ',options:{color:WARM}},{text:'그랜트 20%(①)',options:{color:GREENLT,bold:true}},{text:'·',options:{color:WARM}},{text:'CoC 해제 30% 희석(②)',options:{color:GREENLT,bold:true}},{text:'은 각각 implied 멀티플을 ',options:{color:WARM}},{text:'19.3x→17.7x',options:{color:REDLT,bold:true}},{text:', 둘 다면 ',options:{color:WARM}},{text:'16.6x',options:{color:REDLT,bold:true}},{text:'로 낮춰 peer(8~13x)에 근접시키나 밴드까진 못 미친다. 잔여 갭(~3.6x)은 ',options:{color:WARM}},{text:'ex-AMPC EBITDA를 ~22%로',options:{color:GREENLT,bold:true}},{text:' 올리거나 시장이 프리미엄 멀티플을 줄 때 닫힌다. ★ 순차입금 $6.4bn·부채비중 60% 고정',options:{color:WARM}}],
    {x:MX+0.3,y:6.78,w:W-2*MX-0.6,h:0.78,margin:0,fontFace:BODY,fontSize:9.5,valign:'top',lineSpacingMultiple:1.15});
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
