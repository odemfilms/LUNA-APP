import { useState } from "react";

const MONTHS_DE = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];
const MONTHS_EN = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const WD_DE = ["Mo","Di","Mi","Do","Fr","Sa","So"];
const WD_EN = ["Mo","Tu","We","Th","Fr","Sa","Su"];

function toKey(d) {
  const dt = new Date(d);
  return dt.getFullYear() + "-" + String(dt.getMonth()+1).padStart(2,"0") + "-" + String(dt.getDate()).padStart(2,"0");
}
function addDays(d, n) {
  const r = new Date(d); r.setHours(0,0,0,0); r.setDate(r.getDate()+n); return r;
}
function today0() { const d = new Date(); d.setHours(0,0,0,0); return d; }
function diffDays(a, b) {
  const da = new Date(toKey(a)), db = new Date(toKey(b));
  return Math.round((db - da) / 86400000);
}
function getCycle(lastPeriod, cycleLen) {
  const now = today0();
  const start = new Date(toKey(lastPeriod));
  const daysSince = Math.max(0, diffDays(start, now));
  const cycleDay = (daysSince % cycleLen) + 1;
  const daysLeft = cycleLen - cycleDay + 1;
  const nextPeriod = addDays(now, daysLeft);
  const ovulation = addDays(nextPeriod, -14);
  return {
    cycleDay,
    daysUntilPeriod: daysLeft,
    nextPeriod,
    ovulation,
    fertileStart: addDays(ovulation, -5),
    fertileEnd: addDays(ovulation, 1),
  };
}
function getDayColor(date, periodDays, cycle) {
  const k = toKey(date);
  const tk = toKey(today0());
  if (periodDays[k]) return { bg:"#7B1535", border:"#E8456A", color:"#FFB3C6" };
  if (toKey(cycle.ovulation) === k) return { bg:"#3D3000", border:"#FFD700", color:"#FFE87A" };
  const fk = toKey(date);
  if (fk >= toKey(cycle.fertileStart) && fk <= toKey(cycle.fertileEnd)) return { bg:"#0F3020", border:"#2ECC71", color:"#A8F0C0" };
  if (k === tk) return { bg:"#1A1A3A", border:"#7B8FFF", color:"#C0C8FF" };
  return { bg:"rgba(255,255,255,0.03)", border:"rgba(255,255,255,0.07)", color:"#666" };
}

// ── Embryo SVG sketches ───────────────────────────────────────────────────────
function EmbryoArt({ week }) {
  const skin = "#F9C4CE", dark = "#5C1A30", pink = "#E8456A", light = "#FDDCB5", brown = "#8B6340";
  const glow = "drop-shadow(0 0 12px rgba(232,69,106,0.5))";

  if (week <= 5) return (
    <svg viewBox="0 0 200 200" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g1" cx="40%" cy="35%"><stop offset="0%" stopColor="#F9C4CE"/><stop offset="100%" stopColor="#C97090" stopOpacity="0.6"/></radialGradient>
      </defs>
      <circle cx="100" cy="100" r="45" fill="url(#g1)" stroke="#E8456A" strokeWidth="1.5" strokeDasharray="4 3" opacity="0.7"/>
      {[0,45,90,135,180,225,270,315].map((a,i)=>(
        <circle key={i} cx={100+Math.cos(a*Math.PI/180)*28} cy={100+Math.sin(a*Math.PI/180)*28} r="13" fill="#F9C4CE" stroke="#E8456A" strokeWidth="1" opacity="0.9"/>
      ))}
      <circle cx="100" cy="100" r="10" fill="#E8456A" opacity="0.7"/>
      <circle cx="100" cy="100" r="50" fill="none" stroke="#E8456A" strokeWidth="0.5" opacity="0.2"/>
      <text x="100" y="175" textAnchor="middle" fill="#E8456A" fillOpacity="0.5" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );

  if (week <= 7) return (
    <svg viewBox="0 0 200 220" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g2" cx="40%" cy="35%"><stop offset="0%" stopColor="#F9C4CE"/><stop offset="100%" stopColor="#C97090"/></radialGradient>
      </defs>
      <ellipse cx="100" cy="110" rx="58" ry="70" fill="none" stroke="#E8456A" strokeWidth="0.7" strokeDasharray="3 4" opacity="0.25"/>
      <path d="M100 68 C128 68 145 88 140 112 C135 132 120 146 100 148 C84 148 74 138 76 124" stroke="#E8456A" strokeWidth="20" strokeLinecap="round" fill="none"/>
      <path d="M100 68 C128 68 145 88 140 112 C135 132 120 148 100 148 C84 148 74 138 76 124" stroke={skin} strokeWidth="15" strokeLinecap="round" fill="none"/>
      <circle cx="100" cy="70" r="24" fill="url(#g2)" stroke="#E8456A" strokeWidth="1.5"/>
      <circle cx="93" cy="68" r="3.5" fill={dark} opacity="0.7"/>
      <circle cx="107" cy="68" r="3.5" fill={dark} opacity="0.7"/>
      <circle cx="100" cy="78" r="3" fill="#C97090" opacity="0.5"/>
      <circle cx="114" cy="105" r="6" fill="#E8456A" opacity="0.6"/>
      <text x="100" y="185" textAnchor="middle" fill="#E8456A" fillOpacity="0.4" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );

  if (week <= 9) return (
    <svg viewBox="0 0 200 230" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g3" cx="40%" cy="35%"><stop offset="0%" stopColor="#F9C4CE"/><stop offset="100%" stopColor="#C97090"/></radialGradient>
      </defs>
      <ellipse cx="100" cy="130" rx="22" ry="30" fill="url(#g3)" stroke="#E8456A" strokeWidth="1.5"/>
      <circle cx="100" cy="88" r="30" fill="url(#g3)" stroke="#E8456A" strokeWidth="1.5"/>
      <circle cx="90" cy="85" r="4" fill={dark} opacity="0.6"/>
      <circle cx="110" cy="85" r="4" fill={dark} opacity="0.6"/>
      <ellipse cx="100" cy="96" rx="5" ry="3" fill="#C97090" opacity="0.5"/>
      <ellipse cx="73" cy="88" rx="6" ry="9" fill={skin} stroke="#E8456A" strokeWidth="1" opacity="0.8"/>
      <ellipse cx="127" cy="88" rx="6" ry="9" fill={skin} stroke="#E8456A" strokeWidth="1" opacity="0.8"/>
      <ellipse cx="76" cy="118" rx="9" ry="13" fill={skin} stroke="#E8456A" strokeWidth="1" transform="rotate(-25 76 118)"/>
      <ellipse cx="124" cy="118" rx="9" ry="13" fill={skin} stroke="#E8456A" strokeWidth="1" transform="rotate(25 124 118)"/>
      <ellipse cx="84" cy="155" rx="10" ry="14" fill={skin} stroke="#E8456A" strokeWidth="1" transform="rotate(10 84 155)"/>
      <ellipse cx="116" cy="155" rx="10" ry="14" fill={skin} stroke="#E8456A" strokeWidth="1" transform="rotate(-10 116 155)"/>
      <path d="M100 158 Q108 170 105 180" stroke="#E8456A" strokeWidth="3" strokeLinecap="round"/>
      <text x="100" y="200" textAnchor="middle" fill="#E8456A" fillOpacity="0.4" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );

  if (week <= 12) return (
    <svg viewBox="0 0 200 250" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g4" cx="40%" cy="35%"><stop offset="0%" stopColor="#FDDCB5"/><stop offset="100%" stopColor="#E8A87C" stopOpacity="0.9"/></radialGradient>
      </defs>
      <ellipse cx="100" cy="138" rx="26" ry="36" fill="url(#g4)" stroke="#E8456A" strokeWidth="1.5"/>
      <circle cx="100" cy="90" r="34" fill="url(#g4)" stroke="#E8456A" strokeWidth="1.5"/>
      <ellipse cx="88" cy="87" rx="6" ry="5" fill={dark} opacity="0.55"/>
      <ellipse cx="112" cy="87" rx="6" ry="5" fill={dark} opacity="0.55"/>
      <circle cx="89" cy="86" r="2" fill="white" opacity="0.4"/>
      <circle cx="113" cy="86" r="2" fill="white" opacity="0.4"/>
      <ellipse cx="100" cy="99" rx="5" ry="3" fill="#C97090" opacity="0.5"/>
      <path d="M94 107 Q100 112 106 107" stroke="#C97090" strokeWidth="1.5" fill="none" strokeLinecap="round"/>
      <ellipse cx="69" cy="91" rx="7" ry="10" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <path d="M72 87 Q70 91 72 96" stroke="#E8A87C" strokeWidth="1.2" fill="none"/>
      <ellipse cx="131" cy="91" rx="7" ry="10" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <rect x="93" y="122" width="14" height="14" rx="4" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <path d="M76 130 Q60 140 64 158 Q66 164 74 162" stroke={light} strokeWidth="11" strokeLinecap="round" fill="none"/>
      <path d="M76 130 Q60 140 64 158 Q66 164 74 162" stroke="#E8A87C" strokeWidth="7" strokeLinecap="round" fill="none"/>
      {[-4,-2,0,2,4].map((o,i)=><line key={i} x1={70+o} y1="162" x2={69+o} y2="170" stroke="#C97090" strokeWidth="1.5" strokeLinecap="round"/>)}
      <path d="M124 130 Q140 140 136 158 Q134 164 126 162" stroke={light} strokeWidth="11" strokeLinecap="round" fill="none"/>
      <path d="M124 130 Q140 140 136 158 Q134 164 126 162" stroke="#E8A87C" strokeWidth="7" strokeLinecap="round" fill="none"/>
      {[-4,-2,0,2,4].map((o,i)=><line key={i} x1={130+o} y1="162" x2={131+o} y2="170" stroke="#C97090" strokeWidth="1.5" strokeLinecap="round"/>)}
      <path d="M86 170 Q72 183 76 200 Q80 208 90 204" stroke={light} strokeWidth="12" strokeLinecap="round" fill="none"/>
      <path d="M86 170 Q72 183 76 200 Q80 208 90 204" stroke="#E8A87C" strokeWidth="8" strokeLinecap="round" fill="none"/>
      <path d="M114 170 Q128 183 124 200 Q120 208 110 204" stroke={light} strokeWidth="12" strokeLinecap="round" fill="none"/>
      <path d="M114 170 Q128 183 124 200 Q120 208 110 204" stroke="#E8A87C" strokeWidth="8" strokeLinecap="round" fill="none"/>
      <text x="100" y="228" textAnchor="middle" fill="#E8456A" fillOpacity="0.4" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );

  if (week <= 20) return (
    <svg viewBox="0 0 200 270" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g5" cx="40%" cy="35%"><stop offset="0%" stopColor="#FDDCB5"/><stop offset="100%" stopColor="#D4956A" stopOpacity="0.9"/></radialGradient>
      </defs>
      <ellipse cx="100" cy="148" rx="30" ry="42" fill="url(#g5)" stroke="#E8456A" strokeWidth="1.5"/>
      <circle cx="100" cy="90" r="34" fill="url(#g5)" stroke="#E8456A" strokeWidth="1.5"/>
      <ellipse cx="88" cy="87" rx="7" ry="6" fill={brown} opacity="0.6"/>
      <ellipse cx="112" cy="87" rx="7" ry="6" fill={brown} opacity="0.6"/>
      <circle cx="89" cy="86" r="2.5" fill="white" opacity="0.4"/>
      <circle cx="113" cy="86" r="2.5" fill="white" opacity="0.4"/>
      <path d="M83 83 Q88 79 93 83" stroke={brown} strokeWidth="1.5" fill="none"/>
      <path d="M107 83 Q112 79 117 83" stroke={brown} strokeWidth="1.5" fill="none"/>
      <path d="M96 100 Q100 106 104 100" stroke="#C97090" strokeWidth="1.5" fill="none" strokeLinecap="round"/>
      <circle cx="97" cy="102" r="2" fill="#C97090" opacity="0.4"/>
      <circle cx="103" cy="102" r="2" fill="#C97090" opacity="0.4"/>
      <path d="M92 112 Q96 117 100 115 Q104 117 108 112" stroke="#C97090" strokeWidth="1.5" fill="none" strokeLinecap="round"/>
      <ellipse cx="67" cy="92" rx="7" ry="11" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <path d="M70 87 Q68 92 70 97" stroke="#E8A87C" strokeWidth="1.5" fill="none"/>
      <ellipse cx="133" cy="92" rx="7" ry="11" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <rect x="92" y="122" width="16" height="16" rx="5" fill={light} stroke="#E8456A" strokeWidth="1"/>
      <path d="M72 140 L52 154 L48 178 L58 188" stroke={light} strokeWidth="13" strokeLinecap="round" fill="none"/>
      <path d="M72 140 L52 154 L48 178 L58 188" stroke="#D4956A" strokeWidth="9" strokeLinecap="round" fill="none"/>
      {[-5,-2.5,0,2.5,5].map((o,i)=><line key={i} x1={56+o} y1="188" x2={55+o} y2="198" stroke="#C97090" strokeWidth="2" strokeLinecap="round"/>)}
      <path d="M128 140 L148 154 L152 178 L142 188" stroke={light} strokeWidth="13" strokeLinecap="round" fill="none"/>
      <path d="M128 140 L148 154 L152 178 L142 188" stroke="#D4956A" strokeWidth="9" strokeLinecap="round" fill="none"/>
      {[-5,-2.5,0,2.5,5].map((o,i)=><line key={i} x1={144+o} y1="188" x2={145+o} y2="198" stroke="#C97090" strokeWidth="2" strokeLinecap="round"/>)}
      <path d="M86 188 L76 215 L72 240 L86 246" stroke={light} strokeWidth="14" strokeLinecap="round" fill="none"/>
      <path d="M86 188 L76 215 L72 240 L86 246" stroke="#D4956A" strokeWidth="10" strokeLinecap="round" fill="none"/>
      {[-4,-2,0,2,4].map((o,i)=><line key={i} x1={84+o} y1="246" x2={83+o} y2="255" stroke="#C97090" strokeWidth="2" strokeLinecap="round"/>)}
      <path d="M114 188 L124 215 L128 240 L114 246" stroke={light} strokeWidth="14" strokeLinecap="round" fill="none"/>
      <path d="M114 188 L124 215 L128 240 L114 246" stroke="#D4956A" strokeWidth="10" strokeLinecap="round" fill="none"/>
      {[-4,-2,0,2,4].map((o,i)=><line key={i} x1={116+o} y1="246" x2={117+o} y2="255" stroke="#C97090" strokeWidth="2" strokeLinecap="round"/>)}
      <text x="100" y="265" textAnchor="middle" fill="#E8456A" fillOpacity="0.4" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );

  // Week 21+: Full baby
  const sc = week >= 32 ? 1 : 0.88;
  return (
    <svg viewBox="0 0 240 320" style={{width:"100%",height:"100%",filter:glow}}>
      <defs>
        <radialGradient id="g6" cx="40%" cy="35%"><stop offset="0%" stopColor="#FDDCB5"/><stop offset="100%" stopColor="#D4956A" stopOpacity="0.9"/></radialGradient>
      </defs>
      <ellipse cx="120" cy="178" rx={38*sc} ry={58*sc} fill="url(#g6)" stroke="#E8456A" strokeWidth="2"/>
      <circle cx="120" cy={108*sc+10} r={40*sc} fill="url(#g6)" stroke="#E8456A" strokeWidth="2"/>
      {[...Array(7)].map((_,i)=><path key={i} d={`M ${88+i*8} ${75} Q ${88+i*8} ${66} ${91+i*8} ${69}`} stroke={brown} strokeWidth="2.2" fill="none" strokeLinecap="round"/>)}
      <path d="M 84 82 Q 120 62 156 82" stroke={brown} strokeWidth="3" fill="none" strokeLinecap="round"/>
      <ellipse cx={107*sc+5} cy={100*sc+14} rx="9" ry="8" fill="white" stroke={brown} strokeWidth="1"/>
      <ellipse cx={133*sc+5} cy={100*sc+14} rx="9" ry="8" fill="white" stroke={brown} strokeWidth="1"/>
      <circle cx={108*sc+5} cy={101*sc+14} r="5.5" fill={brown}/>
      <circle cx={134*sc+5} cy={101*sc+14} r="5.5" fill={brown}/>
      <circle cx={109.5*sc+5} cy={99.5*sc+14} r="2" fill="white"/>
      <circle cx={135.5*sc+5} cy={99.5*sc+14} r="2" fill="white"/>
      {[-3,0,3].map((o,i)=><line key={i} x1={107+o} y1="90" x2={106+o} y2="85" stroke={brown} strokeWidth="1.2" strokeLinecap="round"/>)}
      {[-3,0,3].map((o,i)=><line key={i} x1={133+o} y1="90" x2={134+o} y2="85" stroke={brown} strokeWidth="1.2" strokeLinecap="round"/>)}
      <path d="M 98 92 Q 107 88 116 92" stroke={brown} strokeWidth="2" fill="none" strokeLinecap="round"/>
      <path d="M 124 92 Q 133 88 142 92" stroke={brown} strokeWidth="2" fill="none" strokeLinecap="round"/>
      <path d="M 114 112 Q 120 119 126 112" stroke="#C97090" strokeWidth="2" fill="none" strokeLinecap="round"/>
      <circle cx="115" cy="114" r="2.5" fill="#C97090" opacity="0.5"/>
      <circle cx="125" cy="114" r="2.5" fill="#C97090" opacity="0.5"/>
      <ellipse cx="98" cy="118" rx="10" ry="7" fill="#F4A0B0" opacity="0.35"/>
      <ellipse cx="142" cy="118" rx="10" ry="7" fill="#F4A0B0" opacity="0.35"/>
      <path d="M 108 126 Q 113 122 120 124 Q 127 122 132 126" stroke="#C97090" strokeWidth="2" fill="none" strokeLinecap="round"/>
      <path d="M 108 126 Q 120 132 132 126" stroke="#C97090" strokeWidth="1.5" fill="#F4A0B0" fillOpacity="0.3"/>
      <ellipse cx="79" cy="108" rx="8" ry="13" fill={light} stroke="#E8456A" strokeWidth="1.5"/>
      <path d="M 82 103 Q 80 108 82 113" stroke="#E8A87C" strokeWidth="2" fill="none"/>
      <ellipse cx="161" cy="108" rx="8" ry="13" fill={light} stroke="#E8456A" strokeWidth="1.5"/>
      <rect x="111" y="142" width="18" height="18" rx="6" fill={light} stroke="#E8456A" strokeWidth="1.5"/>
      <path d="M 84 162 L 54 178 L 46 206 L 58 218" stroke={light} strokeWidth="18" strokeLinecap="round" fill="none"/>
      <path d="M 84 162 L 54 178 L 46 206 L 58 218" stroke="#D4956A" strokeWidth="12" strokeLinecap="round" fill="none"/>
      {[-6,-3,0,3,6].map((o,i)=><ellipse key={i} cx={56+o} cy={220} rx={2.5} ry={5} fill={light} stroke="#C97090" strokeWidth="1.2"/>)}
      <path d="M 156 162 L 186 178 L 194 206 L 182 218" stroke={light} strokeWidth="18" strokeLinecap="round" fill="none"/>
      <path d="M 156 162 L 186 178 L 194 206 L 182 218" stroke="#D4956A" strokeWidth="12" strokeLinecap="round" fill="none"/>
      {[-6,-3,0,3,6].map((o,i)=><ellipse key={i} cx={184+o} cy={220} rx={2.5} ry={5} fill={light} stroke="#C97090" strokeWidth="1.2"/>)}
      <path d="M 98 233 L 84 266 L 80 294 L 96 302" stroke={light} strokeWidth="19" strokeLinecap="round" fill="none"/>
      <path d="M 98 233 L 84 266 L 80 294 L 96 302" stroke="#D4956A" strokeWidth="13" strokeLinecap="round" fill="none"/>
      {[-5,-2.5,0,2.5,5].map((o,i)=><ellipse key={i} cx={94+o} cy={303} rx={3} ry={5.5} fill={light} stroke="#C97090" strokeWidth="1.2"/>)}
      <path d="M 142 233 L 156 266 L 160 294 L 144 302" stroke={light} strokeWidth="19" strokeLinecap="round" fill="none"/>
      <path d="M 142 233 L 156 266 L 160 294 L 144 302" stroke="#D4956A" strokeWidth="13" strokeLinecap="round" fill="none"/>
      {[-5,-2.5,0,2.5,5].map((o,i)=><ellipse key={i} cx={146+o} cy={303} rx={3} ry={5.5} fill={light} stroke="#C97090" strokeWidth="1.2"/>)}
      <circle cx="120" cy="176" r="5" fill="none" stroke="#E8A87C" strokeWidth="2"/>
      <text x="120" y="315" textAnchor="middle" fill="#E8456A" fillOpacity="0.4" fontSize="11" fontFamily="Georgia,serif">Woche {week}</text>
    </svg>
  );
}

// ── Baby names ────────────────────────────────────────────────────────────────
const NAMES = [
  {name:"Emma",g:"girl",o:"Germanisch",s:0,m:{de:"Ganze, Vollständige",en:"Whole, universal"}},
  {name:"Sophie",g:"girl",o:"Griechisch",s:0,m:{de:"Weisheit",en:"Wisdom"}},
  {name:"Anna",g:"girl",o:"Hebräisch",s:0,m:{de:"Anmut, Gnade",en:"Grace, favour"}},
  {name:"Clara",g:"girl",o:"Lateinisch",s:0,m:{de:"Die Helle, Berühmte",en:"Bright, famous"}},
  {name:"Laura",g:"girl",o:"Lateinisch",s:0,m:{de:"Lorbeerkranz, Siegerin",en:"Laurel, victorious"}},
  {name:"Maria",g:"girl",o:"Hebräisch",s:0,m:{de:"Geliebte",en:"Beloved"}},
  {name:"Julia",g:"girl",o:"Lateinisch",s:0,m:{de:"Die Jugendliche",en:"Youthful"}},
  {name:"Mia",g:"girl",o:"Skandinavisch",s:1,m:{de:"Meine Geliebte",en:"Mine, beloved"}},
  {name:"Lena",g:"girl",o:"Griechisch",s:1,m:{de:"Die Leuchtende",en:"Shining one"}},
  {name:"Zoe",g:"girl",o:"Griechisch",s:1,m:{de:"Das Leben",en:"Life"}},
  {name:"Maya",g:"girl",o:"Sanskrit",s:1,m:{de:"Illusion, Zauber",en:"Illusion, magic"}},
  {name:"Luna",g:"girl",o:"Lateinisch",s:1,m:{de:"Der Mond",en:"The moon"}},
  {name:"Nora",g:"girl",o:"Irisch",s:1,m:{de:"Ehre, Licht",en:"Honor, light"}},
  {name:"Lily",g:"girl",o:"Englisch",s:2,m:{de:"Lilie – Reinheit",en:"Lily, purity"}},
  {name:"Rose",g:"girl",o:"Lateinisch",s:2,m:{de:"Die Rose",en:"Rose, beauty"}},
  {name:"Aurora",g:"girl",o:"Lateinisch",s:2,m:{de:"Morgenröte",en:"Dawn"}},
  {name:"Iris",g:"girl",o:"Griechisch",s:2,m:{de:"Regenbogen",en:"Rainbow"}},
  {name:"Stella",g:"girl",o:"Lateinisch",s:2,m:{de:"Der Stern",en:"Star"}},
  {name:"Freya",g:"girl",o:"Nordisch",s:3,m:{de:"Göttin der Liebe",en:"Goddess of love"}},
  {name:"Victoria",g:"girl",o:"Lateinisch",s:3,m:{de:"Siegerin",en:"Victory"}},
  {name:"Valentina",g:"girl",o:"Lateinisch",s:3,m:{de:"Die Starke",en:"Strong, healthy"}},
  {name:"Amelie",g:"girl",o:"Französisch",s:4,m:{de:"Die Fleißige",en:"Hardworking"}},
  {name:"Mila",g:"girl",o:"Slawisch",s:4,m:{de:"Lieb, Anmutig",en:"Dear, gracious"}},
  {name:"Pia",g:"girl",o:"Lateinisch",s:4,m:{de:"Die Gütige",en:"Pious, kind"}},
  {name:"Amara",g:"girl",o:"Afrikanisch",s:5,m:{de:"Ewige Schönheit",en:"Eternal beauty"}},
  {name:"Sakura",g:"girl",o:"Japanisch",s:5,m:{de:"Kirschblüte",en:"Cherry blossom"}},
  {name:"Layla",g:"girl",o:"Arabisch",s:5,m:{de:"Nacht, Schönheit",en:"Night, beauty"}},
  {name:"Leon",g:"boy",o:"Griechisch",s:0,m:{de:"Löwe, Tapferer",en:"Lion, brave"}},
  {name:"Alexander",g:"boy",o:"Griechisch",s:0,m:{de:"Beschützer der Menschen",en:"Defender of people"}},
  {name:"Felix",g:"boy",o:"Lateinisch",s:0,m:{de:"Der Glückliche",en:"Happy, fortunate"}},
  {name:"Lukas",g:"boy",o:"Griechisch",s:0,m:{de:"Der Leuchtende",en:"Light"}},
  {name:"Sebastian",g:"boy",o:"Griechisch",s:0,m:{de:"Der Ehrwürdige",en:"Venerable"}},
  {name:"Luca",g:"boy",o:"Italienisch",s:1,m:{de:"Der Leuchtende",en:"Bringer of light"}},
  {name:"Noah",g:"boy",o:"Hebräisch",s:1,m:{de:"Ruhe, Trost",en:"Rest, comfort"}},
  {name:"Finn",g:"boy",o:"Irisch",s:1,m:{de:"Hell, Weiß",en:"Fair, bright"}},
  {name:"Leo",g:"boy",o:"Lateinisch",s:1,m:{de:"Löwe",en:"Lion"}},
  {name:"Kai",g:"boy",o:"Hawaiianisch",s:1,m:{de:"Das Meer",en:"The sea"}},
  {name:"Orion",g:"boy",o:"Griechisch",s:2,m:{de:"Jäger am Sternenhimmel",en:"Hunter in the stars"}},
  {name:"River",g:"boy",o:"Englisch",s:2,m:{de:"Fluss",en:"River"}},
  {name:"Sol",g:"boy",o:"Lateinisch",s:2,m:{de:"Die Sonne",en:"The sun"}},
  {name:"Thor",g:"boy",o:"Nordisch",s:3,m:{de:"Donnergott",en:"Thunder god"}},
  {name:"Viktor",g:"boy",o:"Lateinisch",s:3,m:{de:"Sieger",en:"Victor"}},
  {name:"Maximus",g:"boy",o:"Lateinisch",s:3,m:{de:"Der Größte",en:"The greatest"}},
  {name:"Theo",g:"boy",o:"Griechisch",s:4,m:{de:"Gottesgeschenk",en:"Gift of God"}},
  {name:"Emil",g:"boy",o:"Lateinisch",s:4,m:{de:"Der Eifrige",en:"Eager"}},
  {name:"Ole",g:"boy",o:"Nordisch",s:4,m:{de:"Vorfahre",en:"Ancestor"}},
  {name:"Mateo",g:"boy",o:"Spanisch",s:5,m:{de:"Gottesgeschenk",en:"Gift of God"}},
  {name:"Enzo",g:"boy",o:"Italienisch",s:5,m:{de:"Herrscher",en:"Ruler"}},
  {name:"Nico",g:"boy",o:"Griechisch",s:5,m:{de:"Sieger des Volkes",en:"People's victory"}},
];

const EMBRYO_DATA = [
  {week:4, size:"Mohn",mm:"1mm",weight:"<1g",emoji:"🌱",de:"Einnistung abgeschlossen. Winzige Zellkugel.",en:"Implantation complete. Tiny ball of cells."},
  {week:6, size:"Linse",mm:"4mm",weight:"<1g",emoji:"🫛",de:"Herzschlag beginnt. Arm- und Beinknospen.",en:"Heartbeat starts. Arm and leg buds appear."},
  {week:8, size:"Himbeere",mm:"16mm",weight:"1g",emoji:"🍓",de:"Alle Organe vorhanden. Erste Bewegungen.",en:"All organs present. First movements."},
  {week:10,size:"Erdbeere",mm:"31mm",weight:"4g",emoji:"🍓",de:"Embryo wird Fötus. Zehen getrennt.",en:"Embryo becomes fetus. Toes separated."},
  {week:12,size:"Limette",mm:"54mm",weight:"14g",emoji:"🍋",de:"1. Trimester beendet! Reflexe entwickeln sich.",en:"1st trimester done! Reflexes developing."},
  {week:16,size:"Avocado",mm:"116mm",weight:"100g",emoji:"🥑",de:"Augen reagieren auf Licht. Hören beginnt.",en:"Eyes respond to light. Hearing begins."},
  {week:20,size:"Banane",mm:"160mm",weight:"300g",emoji:"🍌",de:"Halbzeit! Erste Tritte spürbar.",en:"Halfway! First kicks felt."},
  {week:24,size:"Maiskolben",mm:"300mm",weight:"600g",emoji:"🌽",de:"Außerhalb überlebensfähig mit Hilfe.",en:"Viable outside with support."},
  {week:28,size:"Aubergine",mm:"380mm",weight:"1kg",emoji:"🍆",de:"3. Trimester! Augen öffnen sich.",en:"3rd trimester! Eyes open."},
  {week:32,size:"Ananas",mm:"420mm",weight:"1.7kg",emoji:"🍍",de:"Lunge fast fertig. Schlaf-Wach-Rhythmus.",en:"Lungs nearly ready. Sleep-wake cycle."},
  {week:36,size:"Papaya",mm:"470mm",weight:"2.6kg",emoji:"🥭",de:"Fast fertig! Kopf dreht sich nach unten.",en:"Almost ready! Head turning down."},
  {week:40,size:"Wassermelone",mm:"510mm",weight:"3.4kg",emoji:"🍉",de:"Geburtstermin! Baby ist voll ausgewachsen.",en:"Due date! Baby fully grown."},
];

const STYLES = {de:["Klassisch","Modern","Natur","Stark","Süß","International"],en:["Classic","Modern","Nature","Strong","Sweet","International"]};

// ── Main App ──────────────────────────────────────────────────────────────────
export default function LunaApp() {
  const [lang, setLang] = useState("de");
  const [tab, setTab] = useState(0);
  const [curMonth, setCurMonth] = useState(new Date());
  const [cycleLen, setCycleLen] = useState(28);
  const [periodLen, setPeriodLen] = useState(5);
  const [lastPeriod, setLastPeriod] = useState(() => {
    const d = new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate()-14); return d;
  });
  const [periodDays, setPeriodDays] = useState({});
  const [sexLog, setSexLog] = useState({});
  const [pregMode, setPregMode] = useState(false);
  const [lmp, setLmp] = useState(() => {
    const d = new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate()-70); return d;
  });
  const [sexModal, setSexModal] = useState(false);
  const [selDay, setSelDay] = useState(null);
  const [nameSearch, setNameSearch] = useState("");
  const [nameG, setNameG] = useState("both");
  const [nameS, setNameS] = useState(-1);
  const [favs, setFavs] = useState([]);
  const [showFavs, setShowFavs] = useState(false);
  const [selName, setSelName] = useState(null);

  const now = today0();
  const cycle = getCycle(lastPeriod, cycleLen);
  const weeksP = Math.max(4, Math.floor(diffDays(new Date(toKey(lmp)), now) / 7));
  const embryo = EMBRYO_DATA.reduce((p,c)=>c.week<=weeksP?c:p, EMBRYO_DATA[0]);
  const months = lang==="de" ? MONTHS_DE : MONTHS_EN;
  const wd = lang==="de" ? WD_DE : WD_EN;
  const styles = STYLES[lang];

  function togglePeriod(date) {
    const k = toKey(date);
    setPeriodDays(p => { const n={...p}; if(n[k]) delete n[k]; else n[k]=true; return n; });
  }
  function logSex(type) {
    setSexLog(s => ({...s,[toKey(selDay||now)]:type}));
    setSexModal(false);
  }
  function toggleFav(name) {
    setFavs(f => f.includes(name)?f.filter(n=>n!==name):[...f,name]);
  }

  // Calendar
  function buildCal() {
    const y=curMonth.getFullYear(), m=curMonth.getMonth();
    let dow = new Date(y,m,1).getDay(); dow = dow===0?6:dow-1;
    const cells = [];
    for(let i=0;i<dow;i++) cells.push(null);
    for(let d=1;d<=new Date(y,m+1,0).getDate();d++) cells.push(new Date(y,m,d));
    return cells;
  }
  const cells = buildCal();

  const filtNames = NAMES.filter(n => {
    if(nameG!=="both"&&n.g!==nameG) return false;
    if(nameS>=0&&n.s!==nameS) return false;
    if(nameSearch&&!n.name.toLowerCase().includes(nameSearch.toLowerCase())) return false;
    if(showFavs&&!favs.includes(n.name)) return false;
    return true;
  });

  const C = "#E8456A";
  const card = {background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:16,padding:16};

  // Mindmap home
  const nodes = [
    {i:1,icon:"🗓",label:lang==="de"?"Kalender":"Calendar",color:"#E8456A",x:"50%",y:"8%"},
    {i:2,icon:"📍",label:lang==="de"?"Heute":"Today",color:"#FF85A1",x:"83%",y:"32%"},
    {i:3,icon:"🌟",label:lang==="de"?"Eisprung":"Ovulation",color:"#FFD700",x:"74%",y:"72%"},
    {i:4,icon:"👶",label:lang==="de"?"Baby":"Baby",color:"#C97090",x:"26%",y:"72%"},
    {i:5,icon:"💝",label:lang==="de"?"Namen":"Names",color:"#FF85A1",x:"17%",y:"32%"},
    {i:6,icon:"⚙️",label:lang==="de"?"Settings":"Settings",color:"#7B8FFF",x:"50%",y:"88%"},
  ];

  return (
    <div style={{fontFamily:"'Cormorant Garamond',Georgia,serif",background:"#08040F",minHeight:"100vh",color:"white"}}>
      <div style={{position:"fixed",inset:0,pointerEvents:"none",zIndex:0}}>
        <div style={{position:"absolute",inset:0,background:"radial-gradient(ellipse at 25% 15%, rgba(61,10,32,0.3) 0%, transparent 55%)"}}/>
        <div style={{position:"absolute",inset:0,background:"radial-gradient(ellipse at 75% 85%, rgba(10,16,32,0.3) 0%, transparent 55%)"}}/>
      </div>

      <div style={{position:"relative",zIndex:10,maxWidth:480,margin:"0 auto",padding:"16px 16px 90px"}}>
        {/* Header */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div onClick={()=>setTab(0)} style={{cursor:"pointer"}}>
            <div style={{fontSize:10,letterSpacing:"0.3em",color:C,opacity:0.6}}>◈ {lang==="de"?"Dein Zyklus-Begleiter":"Your Cycle Companion"}</div>
            <div style={{fontSize:28,fontWeight:700,letterSpacing:"0.12em",textShadow:"0 0 30px rgba(232,69,106,0.4)"}}>🌙 LUNA</div>
          </div>
          <button onClick={()=>setLang(l=>l==="de"?"en":"de")} style={{background:"rgba(232,69,106,0.1)",border:"1px solid rgba(232,69,106,0.3)",color:C,borderRadius:8,padding:"5px 11px",fontSize:11,cursor:"pointer"}}>{lang==="de"?"EN":"DE"}</button>
        </div>

        {/* ── HOME Mindmap ── */}
        {tab===0 && (
          <div>
            <div style={{textAlign:"center",marginBottom:20}}>
              <div style={{fontSize:20,fontWeight:700,letterSpacing:"0.1em"}}>{lang==="de"?"Willkommen bei LUNA":"Welcome to LUNA"}</div>
              <div style={{fontSize:12,color:"#555",marginTop:4}}>{lang==="de"?"Wähle einen Bereich":"Choose a section"}</div>
            </div>
            <div style={{position:"relative",height:340}}>
              <svg style={{position:"absolute",inset:0,width:"100%",height:"100%"}}>
                {nodes.map(n=>(
                  <line key={n.i} x1="50%" y1="50%" x2={n.x} y2={n.y} stroke={n.color} strokeWidth="1" strokeOpacity="0.25" strokeDasharray="4 4"/>
                ))}
              </svg>
              <div style={{position:"absolute",left:"50%",top:"50%",transform:"translate(-50%,-50%)",width:60,height:60,borderRadius:"50%",background:"rgba(232,69,106,0.12)",border:"2px solid rgba(232,69,106,0.4)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,boxShadow:"0 0 24px rgba(232,69,106,0.25)"}}>🌙</div>
              {nodes.map(n=>(
                <div key={n.i} onClick={()=>setTab(n.i)}
                  style={{position:"absolute",left:n.x,top:n.y,transform:"translate(-50%,-50%)",width:66,height:66,borderRadius:"50%",background:`${n.color}14`,border:`2px solid ${n.color}44`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",transition:"all 0.2s",boxShadow:"none",zIndex:2}}>
                  <span style={{fontSize:22}}>{n.icon}</span>
                  <span style={{fontSize:8,color:n.color,marginTop:2,textAlign:"center",fontWeight:600,maxWidth:54,lineHeight:1.2}}>{n.label}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── KALENDER ── */}
        {tab===1 && (
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <button onClick={()=>setCurMonth(m=>{const n=new Date(m);n.setMonth(n.getMonth()-1);return n;})} style={{background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.1)",color:"white",borderRadius:10,padding:"6px 14px",cursor:"pointer",fontSize:16}}>←</button>
              <div style={{fontSize:16,fontWeight:700}}>{months[curMonth.getMonth()]} {curMonth.getFullYear()}</div>
              <button onClick={()=>setCurMonth(m=>{const n=new Date(m);n.setMonth(n.getMonth()+1);return n;})} style={{background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.1)",color:"white",borderRadius:10,padding:"6px 14px",cursor:"pointer",fontSize:16}}>→</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:4}}>
              {wd.map(d=><div key={d} style={{textAlign:"center",fontSize:10,color:"#444",padding:"2px 0"}}>{d}</div>)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
              {cells.map((date,i)=>{
                if(!date) return <div key={i}/>;
                const col = getDayColor(date, periodDays, cycle);
                const isSelected = selDay && toKey(date)===toKey(selDay);
                return (
                  <div key={i} onClick={()=>setSelDay(date)}
                    style={{aspectRatio:"1",borderRadius:10,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",fontSize:12,fontWeight:toKey(date)===toKey(now)?700:400,background:col.bg,border:`1.5px solid ${isSelected?"white":col.border}`,color:col.color,transition:"all 0.1s"}}>
                    <span>{date.getDate()}</span>
                    {sexLog[toKey(date)] && <div style={{fontSize:7}}>{sexLog[toKey(date)]==="protected"?"🛡":"❤️"}</div>}
                    {toKey(date)===toKey(cycle.ovulation)&&!periodDays[toKey(date)]&&<div style={{fontSize:7}}>★</div>}
                  </div>
                );
              })}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:10,marginTop:12,justifyContent:"center"}}>
              {[["🔴",lang==="de"?"Periode":"Period"],["🟢",lang==="de"?"Fruchtbar":"Fertile"],["⭐",lang==="de"?"Eisprung":"Ovulation"],["🔵",lang==="de"?"Heute":"Today"]].map(([ic,lb])=>(
                <div key={lb} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:"#555"}}><span>{ic}</span><span>{lb}</span></div>
              ))}
            </div>
            {selDay && (
              <div style={{marginTop:10,textAlign:"center",fontSize:11,color:C,background:"rgba(232,69,106,0.08)",border:"1px solid rgba(232,69,106,0.2)",borderRadius:10,padding:7}}>
                📅 {selDay.toLocaleDateString(lang==="de"?"de-DE":"en-GB",{weekday:"long",day:"numeric",month:"long"})}
              </div>
            )}
            <div style={{display:"flex",gap:10,marginTop:10}}>
              <button onClick={()=>togglePeriod(selDay||now)} style={{flex:1,padding:"11px 8px",borderRadius:12,background:"rgba(232,69,106,0.12)",border:"1px solid rgba(232,69,106,0.35)",color:C,cursor:"pointer",fontSize:12,fontWeight:700}}>
                🩸 {periodDays[toKey(selDay||now)] ? (lang==="de"?"Periode beenden":"End Period") : (lang==="de"?"Periode starten":"Start Period")}
              </button>
              <button onClick={()=>setSexModal(true)} style={{flex:1,padding:"11px 8px",borderRadius:12,background:"rgba(255,150,180,0.07)",border:"1px solid rgba(255,150,180,0.25)",color:"#FFB3C6",cursor:"pointer",fontSize:12,fontWeight:700}}>
                💕 {lang==="de"?"Intimität loggen":"Log Intimacy"}
              </button>
            </div>
          </div>
        )}

        {/* ── HEUTE ── */}
        {tab===2 && (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{...card,background:"rgba(232,69,106,0.07)",border:"1px solid rgba(232,69,106,0.18)",textAlign:"center",padding:22}}>
              <div style={{fontSize:10,letterSpacing:"0.3em",color:C,opacity:0.7,marginBottom:5}}>◈ {lang==="de"?"Zyklustag":"Cycle Day"} ◈</div>
              <div style={{fontSize:72,fontWeight:700,color:C,textShadow:"0 0 40px rgba(232,69,106,0.55)",lineHeight:1}}>{cycle.cycleDay}</div>
              <div style={{fontSize:12,color:"#666",marginTop:4}}>/ {cycleLen} {lang==="de"?"Tage":"days"}</div>
              <svg viewBox="0 0 120 120" style={{width:100,height:100,margin:"10px auto 0"}}>
                <circle cx="60" cy="60" r="50" fill="none" stroke="rgba(232,69,106,0.1)" strokeWidth="8"/>
                <circle cx="60" cy="60" r="50" fill="none" stroke={C} strokeWidth="8" strokeLinecap="round" strokeDasharray={`${(cycle.cycleDay/cycleLen)*314} 314`} transform="rotate(-90 60 60)" style={{filter:"drop-shadow(0 0 5px #E8456A)"}}/>
              </svg>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[{label:lang==="de"?"Nächste Periode":"Next Period",days:cycle.daysUntilPeriod,color:C,icon:"🩸"},
                {label:lang==="de"?"Eisprung":"Ovulation",days:diffDays(now,cycle.ovulation),color:"#FFD700",icon:"⭐"}].map(({label,days,color,icon})=>(
                <div key={label} style={{...card,background:`${color}10`,border:`1px solid ${color}25`,textAlign:"center",padding:14}}>
                  <div style={{fontSize:22,marginBottom:3}}>{icon}</div>
                  <div style={{fontSize:10,color:"#666",marginBottom:3}}>{label}</div>
                  <div style={{fontSize:26,fontWeight:700,color}}>{Math.max(0,days)}</div>
                  <div style={{fontSize:10,color:"#555"}}>{lang==="de"?"Tage":"days"}</div>
                </div>
              ))}
            </div>
            <div style={{...card,background:"rgba(46,204,113,0.06)",border:"1px solid rgba(46,204,113,0.16)",padding:14}}>
              <div style={{fontSize:10,letterSpacing:"0.18em",color:"#2ECC71",marginBottom:8}}>◈ {lang==="de"?"Fruchtbares Fenster":"Fertile Window"}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div><div style={{fontSize:12,color:"#888"}}>{cycle.fertileStart.toLocaleDateString(lang==="de"?"de-DE":"en-GB",{day:"numeric",month:"short"})}</div><div style={{fontSize:10,color:"#444"}}>Start</div></div>
                <div style={{color:"#2ECC71",fontSize:18}}>→</div>
                <div style={{textAlign:"right"}}><div style={{fontSize:12,color:"#888"}}>{cycle.fertileEnd.toLocaleDateString(lang==="de"?"de-DE":"en-GB",{day:"numeric",month:"short"})}</div><div style={{fontSize:10,color:"#444"}}>End</div></div>
                <div><div style={{fontSize:22,color:"#FFD700"}}>⭐</div><div style={{fontSize:11,color:"#888"}}>{cycle.ovulation.toLocaleDateString(lang==="de"?"de-DE":"en-GB",{day:"numeric",month:"short"})}</div></div>
              </div>
            </div>
            <div style={{...card,background:"rgba(255,150,180,0.05)",border:"1px solid rgba(255,150,180,0.15)",padding:14}}>
              <div style={{fontSize:10,letterSpacing:"0.18em",color:"#FFB3C6",marginBottom:8}}>◈ {lang==="de"?"Intimität loggen":"Log Intimacy"}</div>
              <div style={{display:"flex",gap:10}}>
                {[["protected","🛡️",lang==="de"?"Geschützt":"Protected","#7B8FFF"],["unprotected","❤️",lang==="de"?"Ungeschützt":"Unprotected",C]].map(([type,icon,label,color])=>(
                  <button key={type} onClick={()=>{setSelDay(now);logSex(type);}} style={{flex:1,padding:"10px 6px",borderRadius:12,cursor:"pointer",fontSize:11,fontWeight:sexLog[toKey(now)]===type?700:400,background:sexLog[toKey(now)]===type?`${color}20`:"rgba(255,255,255,0.04)",border:`1.5px solid ${sexLog[toKey(now)]===type?color:"rgba(255,255,255,0.1)"}`,color:sexLog[toKey(now)]===type?color:"#666",transition:"all 0.2s"}}>{icon} {label}</button>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── EISPRUNG ── */}
        {tab===3 && (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{...card,background:"rgba(255,215,0,0.06)",border:"1px solid rgba(255,215,0,0.16)",textAlign:"center",padding:22}}>
              <div style={{fontSize:10,letterSpacing:"0.28em",color:"#FFD700",marginBottom:8}}>◈ {lang==="de"?"Eisprung":"Ovulation"} ◈</div>
              <div style={{fontSize:40,marginBottom:6}}>⭐</div>
              <div style={{fontSize:18,fontWeight:700,color:"#FFD700"}}>{cycle.ovulation.toLocaleDateString(lang==="de"?"de-DE":"en-GB",{weekday:"long",day:"numeric",month:"long"})}</div>
              <div style={{fontSize:12,color:"#666",marginTop:5}}>{lang==="de"?"In":"In"} {Math.max(0,diffDays(now,cycle.ovulation))} {lang==="de"?"Tagen":"days"}</div>
            </div>
            <div style={card}>
              <div style={{fontSize:10,color:"#555",letterSpacing:"0.14em",marginBottom:12}}>{lang==="de"?"Fruchtbarkeitsfenster":"Fertility Window"}</div>
              <div style={{display:"flex",gap:5,overflowX:"auto",paddingBottom:4}}>
                {[-6,-5,-4,-3,-2,-1,0,1].map(offset=>{
                  const d=addDays(cycle.ovulation,offset);
                  const isOv=offset===0, isFert=offset>=-5&&offset<=0;
                  return (
                    <div key={offset} style={{flex:"0 0 auto",width:48,borderRadius:11,padding:"8px 3px",textAlign:"center",background:isOv?"rgba(255,215,0,0.16)":isFert?"rgba(46,204,113,0.08)":"rgba(255,255,255,0.025)",border:`1.5px solid ${isOv?"#FFD700":isFert?"rgba(46,204,113,0.35)":"rgba(255,255,255,0.07)"}`,boxShadow:isOv?"0 0 12px rgba(255,215,0,0.25)":"none"}}>
                      <div style={{fontSize:9,color:"#444",marginBottom:2}}>{wd[(d.getDay()+6)%7]}</div>
                      <div style={{fontSize:13,fontWeight:isOv?700:400,color:isOv?"#FFD700":isFert?"#2ECC71":"#555"}}>{d.getDate()}</div>
                      <div style={{fontSize:9,color:"#444",marginTop:1}}>{months[d.getMonth()].slice(0,3)}</div>
                      {isOv&&<div style={{fontSize:9,marginTop:2}}>⭐</div>}
                    </div>
                  );
                })}
              </div>
            </div>
            <div style={card}>
              {[{label:lang==="de"?"Zykluslänge":"Cycle Length",val:cycleLen,set:setCycleLen,min:21,max:40},
                {label:lang==="de"?"Periodendauer":"Period Length",val:periodLen,set:setPeriodLen,min:2,max:10}].map(({label,val,set,min,max})=>(
                <div key={label} style={{marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><span style={{fontSize:12,color:"#888"}}>{label}</span><span style={{fontSize:12,color:C,fontWeight:700}}>{val} {lang==="de"?"Tage":"days"}</span></div>
                  <input type="range" min={min} max={max} value={val} onChange={e=>set(+e.target.value)} style={{width:"100%",accentColor:C}}/>
                </div>
              ))}
              <div><div style={{fontSize:12,color:"#888",marginBottom:6}}>{lang==="de"?"Letzte Periode":"Last Period"}</div>
                <input type="date" value={toKey(lastPeriod)} onChange={e=>{const d=new Date(e.target.value);d.setHours(0,0,0,0);setLastPeriod(d);}} style={{width:"100%",background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:9,padding:"9px 11px",color:"white",fontSize:13}}/>
              </div>
            </div>
          </div>
        )}

        {/* ── BABY ── */}
        {tab===4 && (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",...card,padding:14}}>
              <span style={{fontSize:12,color:"#888"}}>{lang==="de"?"Schwangerschaftsmodus":"Pregnancy Mode"}</span>
              <button onClick={()=>setPregMode(p=>!p)} style={{width:48,height:26,borderRadius:13,border:"none",cursor:"pointer",background:pregMode?"linear-gradient(90deg,#E8456A,#FF85A1)":"rgba(255,255,255,0.1)",position:"relative",transition:"all 0.3s"}}>
                <div style={{position:"absolute",top:3,width:20,height:20,borderRadius:"50%",background:"white",transition:"left 0.3s",left:pregMode?25:2,boxShadow:"0 2px 4px rgba(0,0,0,0.3)"}}/>
              </button>
            </div>
            {pregMode ? (
              <>
                <div style={{...card,padding:14}}>
                  <div style={{fontSize:10,letterSpacing:"0.15em",color:C,marginBottom:7}}>◈ {lang==="de"?"LMP Datum (letzte Periode)":"LMP Date (last period)"}</div>
                  <input type="date" value={toKey(lmp)} onChange={e=>{const d=new Date(e.target.value);d.setHours(0,0,0,0);setLmp(d);}} style={{width:"100%",background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:9,padding:"9px 11px",color:"white",fontSize:13}}/>
                </div>
                <div style={{...card,background:"rgba(232,69,106,0.06)",border:"1px solid rgba(232,69,106,0.2)",textAlign:"center",padding:20}}>
                  <div style={{fontSize:10,letterSpacing:"0.28em",color:C,marginBottom:5}}>◈ {lang==="de"?"Schwangerschaftswoche":"Pregnancy Week"} ◈</div>
                  <div style={{fontSize:66,fontWeight:700,color:C,textShadow:"0 0 40px rgba(232,69,106,0.5)",lineHeight:1}}>{weeksP}</div>
                  <div style={{fontSize:13,color:"#666",marginTop:3}}>{lang==="de"?"Woche":"Week"}</div>
                  <div style={{marginTop:8,fontSize:11,color:"#555"}}>{lang==="de"?"Voraus. Geburtstermin":"Est. Due Date"}: <span style={{color:"#FFB3C6"}}>{addDays(new Date(toKey(lmp)),280).toLocaleDateString(lang==="de"?"de-DE":"en-GB",{day:"numeric",month:"long",year:"numeric"})}</span></div>
                  <div style={{fontSize:10,color:"#444",marginTop:3}}>{weeksP<=13?(lang==="de"?"1. Trimester":"1st Trimester"):weeksP<=26?(lang==="de"?"2. Trimester":"2nd Trimester"):(lang==="de"?"3. Trimester":"3rd Trimester")}</div>
                </div>
                <div style={{...card,border:"1px solid rgba(255,150,180,0.15)",padding:18}}>
                  <div style={{fontSize:10,letterSpacing:"0.15em",color:"#FFB3C6",marginBottom:14,textAlign:"center"}}>◈ {lang==="de"?"Dein Baby · Woche":"Your Baby · Week"} {weeksP} ◈</div>
                  <div style={{width:"100%",maxWidth:200,height:220,margin:"0 auto"}}>
                    <EmbryoArt week={weeksP}/>
                  </div>
                  <div style={{textAlign:"center",marginTop:14}}>
                    <div style={{fontSize:26,marginBottom:6}}>{embryo.emoji}</div>
                    <div style={{fontSize:16,fontWeight:700,color:"#FFB3C6",marginBottom:3}}>{embryo.size}</div>
                    <div style={{fontSize:12,color:"#666",marginBottom:3}}>{embryo.mm} · {lang==="de"?"Gewicht ca.":"Weight approx."} {embryo.weight}</div>
                    <div style={{fontSize:12,color:"#999",lineHeight:1.6,fontStyle:"italic"}}>{lang==="de"?embryo.de:embryo.en}</div>
                  </div>
                </div>
                <div style={card}>
                  <div style={{fontSize:10,color:"#555",letterSpacing:"0.14em",marginBottom:12}}>{lang==="de"?"Entwicklungs-Timeline":"Development Timeline"}</div>
                  {EMBRYO_DATA.map(w=>(
                    <div key={w.week} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderRadius:11,marginBottom:5,background:w.week===embryo.week?"rgba(232,69,106,0.1)":"rgba(255,255,255,0.02)",border:`1px solid ${w.week===embryo.week?"rgba(232,69,106,0.32)":"rgba(255,255,255,0.05)"}`,opacity:w.week>weeksP?0.3:1}}>
                      <span style={{fontSize:16}}>{w.emoji}</span>
                      <div style={{flex:1}}><div style={{fontSize:11,fontWeight:700,color:w.week===embryo.week?C:"#888"}}>{lang==="de"?"Woche":"Week"} {w.week}</div><div style={{fontSize:10,color:"#444"}}>{w.size} · {w.mm}</div></div>
                      {w.week<=weeksP&&<span style={{color:"#2ECC71",fontSize:10}}>✓</span>}
                      {w.week===embryo.week&&<span style={{color:C,fontSize:8,fontWeight:700}}>● NOW</span>}
                    </div>
                  ))}
                </div>
              </>
            ):(
              <div style={{textAlign:"center",padding:"50px 20px",color:"#333"}}>
                <div style={{fontSize:40,marginBottom:12}}>🌙</div>
                <div style={{fontSize:14,color:"#444",lineHeight:1.6}}>{lang==="de"?"Aktiviere den Schwangerschaftsmodus oben.":"Enable pregnancy mode above."}</div>
              </div>
            )}
          </div>
        )}

        {/* ── NAMEN ── */}
        {tab===5 && (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={{fontSize:10,letterSpacing:"0.28em",color:C,textAlign:"center"}}>◈ {lang==="de"?"Namensfinder":"Name Finder"} ◈</div>
            <input value={nameSearch} onChange={e=>setNameSearch(e.target.value)} placeholder={lang==="de"?"Namen suchen...":"Search names..."} style={{width:"100%",background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:11,padding:"10px 13px",color:"white",fontSize:14,boxSizing:"border-box"}}/>
            <div style={{display:"flex",gap:7}}>
              {[["both",lang==="de"?"Alle":"All",C],["girl","♀ "+( lang==="de"?"Mädchen":"Girl"),"#FF85A1"],["boy","♂ "+(lang==="de"?"Junge":"Boy"),"#7B8FFF"]].map(([val,label,color])=>(
                <button key={val} onClick={()=>setNameG(val)} style={{flex:1,padding:"8px 4px",borderRadius:11,cursor:"pointer",fontSize:11,fontWeight:nameG===val?700:400,background:nameG===val?`${color}20`:"rgba(255,255,255,0.04)",border:`1.5px solid ${nameG===val?color:"rgba(255,255,255,0.1)"}`,color:nameG===val?color:"#555",transition:"all 0.2s"}}>{label}</button>
              ))}
            </div>
            <div style={{display:"flex",gap:5,overflowX:"auto",paddingBottom:3}}>
              <button onClick={()=>setNameS(-1)} style={{flex:"0 0 auto",padding:"6px 11px",borderRadius:9,cursor:"pointer",fontSize:11,background:nameS===-1?"rgba(232,69,106,0.16)":"rgba(255,255,255,0.04)",border:`1px solid ${nameS===-1?C:"rgba(255,255,255,0.09)"}`,color:nameS===-1?C:"#555"}}>{lang==="de"?"Alle":"All"}</button>
              {styles.map((s,i)=>(
                <button key={i} onClick={()=>setNameS(i)} style={{flex:"0 0 auto",padding:"6px 11px",borderRadius:9,cursor:"pointer",fontSize:11,background:nameS===i?"rgba(232,69,106,0.16)":"rgba(255,255,255,0.04)",border:`1px solid ${nameS===i?C:"rgba(255,255,255,0.09)"}`,color:nameS===i?C:"#555"}}>{s}</button>
              ))}
            </div>
            <button onClick={()=>setShowFavs(f=>!f)} style={{padding:"8px 13px",borderRadius:10,cursor:"pointer",fontSize:11,fontWeight:showFavs?700:400,background:showFavs?"rgba(255,215,0,0.13)":"rgba(255,255,255,0.04)",border:`1px solid ${showFavs?"#FFD700":"rgba(255,255,255,0.09)"}`,color:showFavs?"#FFD700":"#555"}}>⭐ {lang==="de"?"Favoriten":"Favorites"} {favs.length>0&&`(${favs.length})`}</button>
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {filtNames.length===0?(<div style={{textAlign:"center",padding:"28px 0",color:"#333",fontSize:13}}>{lang==="de"?"Keine Treffer":"No results"}</div>):
                filtNames.map(n=>{
                  const isFav=favs.includes(n.name), isSel=selName?.name===n.name;
                  return (
                    <div key={n.name}>
                      <div onClick={()=>setSelName(isSel?null:n)} style={{display:"flex",alignItems:"center",gap:11,padding:"12px 13px",borderRadius:13,cursor:"pointer",background:isSel?"rgba(232,69,106,0.1)":"rgba(255,255,255,0.03)",border:`1px solid ${isSel?"rgba(232,69,106,0.38)":"rgba(255,255,255,0.06)"}`,transition:"all 0.15s"}}>
                        <div style={{flex:1}}>
                          <div style={{display:"flex",alignItems:"center",gap:7,flexWrap:"wrap"}}>
                            <span style={{fontSize:17,fontWeight:700,color:isSel?C:"white"}}>{n.name}</span>
                            <span style={{fontSize:9,color:n.g==="girl"?"#FF85A1":"#7B8FFF",background:n.g==="girl"?"rgba(255,133,161,0.1)":"rgba(123,143,255,0.1)",padding:"2px 6px",borderRadius:5}}>{n.g==="girl"?"♀":"♂"}</span>
                            <span style={{fontSize:9,color:"#444",background:"rgba(255,255,255,0.04)",padding:"2px 6px",borderRadius:5}}>{styles[n.s]}</span>
                          </div>
                          <div style={{fontSize:10,color:"#444",marginTop:2}}>{n.o}</div>
                        </div>
                        <button onClick={e=>{e.stopPropagation();toggleFav(n.name);}} style={{background:isFav?"rgba(255,215,0,0.15)":"rgba(255,255,255,0.04)",border:`1px solid ${isFav?"#FFD700":"rgba(255,255,255,0.09)"}`,borderRadius:8,padding:"5px 9px",cursor:"pointer",fontSize:13,color:isFav?"#FFD700":"#444",transition:"all 0.15s"}}>{isFav?"⭐":"☆"}</button>
                      </div>
                      {isSel&&(
                        <div style={{background:"rgba(232,69,106,0.05)",border:"1px solid rgba(232,69,106,0.18)",borderRadius:11,padding:"10px 13px",marginTop:3,fontSize:12,color:"#999",fontStyle:"italic",lineHeight:1.6}}>
                          <span style={{color:C,fontSize:9,letterSpacing:"0.14em",fontStyle:"normal"}}>◈ {lang==="de"?"Bedeutung":"Meaning"}: </span>{n.m[lang]}
                        </div>
                      )}
                    </div>
                  );
                })}
            </div>
          </div>
        )}

        {/* ── EINSTELLUNGEN ── */}
        {tab===6 && (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <div style={card}>
              <div style={{fontSize:10,letterSpacing:"0.15em",color:"#555",marginBottom:12}}>{lang==="de"?"Sprache":"Language"}</div>
              <div style={{display:"flex",gap:10}}>
                {["de","en"].map(l=>(
                  <button key={l} onClick={()=>setLang(l)} style={{flex:1,padding:11,borderRadius:11,cursor:"pointer",fontWeight:700,fontSize:13,background:lang===l?"rgba(232,69,106,0.16)":"rgba(255,255,255,0.04)",border:`1.5px solid ${lang===l?C:"rgba(255,255,255,0.09)"}`,color:lang===l?C:"#555"}}>
                    {l==="de"?"🇩🇪 Deutsch":"🇬🇧 English"}
                  </button>
                ))}
              </div>
            </div>
            <div style={card}>
              <div style={{fontSize:13,color:"#444",lineHeight:1.7}}>{lang==="de"?"LUNA ist dein privater Zyklus-Begleiter. Alle Daten bleiben auf deinem Gerät. 🌙":"LUNA is your private cycle companion. All data stays on your device. 🌙"}</div>
              <div style={{marginTop:10,fontSize:10,color:"#2a2a2a"}}>v4.0 · LUNA © 2025</div>
            </div>
          </div>
        )}
      </div>

      {/* Bottom nav */}
      {tab>0 && (
        <div style={{position:"fixed",bottom:0,left:0,right:0,zIndex:50,background:"rgba(8,4,15,0.97)",borderTop:"1px solid rgba(232,69,106,0.12)",padding:"8px 6px 12px",display:"flex",gap:3,maxWidth:480,margin:"0 auto"}}>
          <button onClick={()=>setTab(0)} style={{flex:1,padding:"6px 2px",background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:9,color:"#555",cursor:"pointer",fontSize:9}}>🌙</button>
          {[1,2,3,4,5,6].map(i=>(
            <button key={i} onClick={()=>setTab(i)} style={{flex:1,padding:"6px 2px",background:tab===i?"rgba(232,69,106,0.15)":"rgba(255,255,255,0.03)",border:`1px solid ${tab===i?"rgba(232,69,106,0.4)":"rgba(255,255,255,0.06)"}`,borderRadius:9,color:tab===i?C:"#555",cursor:"pointer",fontSize:14}}>
              {["🗓","📍","🌟","👶","💝","⚙️"][i-1]}
            </button>
          ))}
        </div>
      )}

      {/* Sex modal */}
      {sexModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:100,display:"flex",alignItems:"flex-end"}} onClick={()=>setSexModal(false)}>
          <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxWidth:480,margin:"0 auto",background:"#0F0816",border:"1px solid rgba(232,69,106,0.25)",borderRadius:"20px 20px 0 0",padding:24}}>
            <div style={{textAlign:"center",fontSize:10,letterSpacing:"0.28em",color:C,marginBottom:16}}>◈ {lang==="de"?"Intimität loggen":"Log Intimacy"} ◈</div>
            <div style={{display:"flex",gap:12}}>
              {[["protected","🛡️",lang==="de"?"Geschützt":"Protected","#7B8FFF"],["unprotected","❤️",lang==="de"?"Ungeschützt":"Unprotected",C]].map(([type,icon,label,color])=>(
                <button key={type} onClick={()=>logSex(type)} style={{flex:1,padding:16,borderRadius:14,cursor:"pointer",fontSize:13,fontWeight:700,background:`${color}14`,border:`1.5px solid ${color}50`,color}}>
                  <div style={{fontSize:24,marginBottom:6}}>{icon}</div>{label}
                </button>
              ))}
            </div>
            <button onClick={()=>setSexModal(false)} style={{width:"100%",marginTop:10,padding:12,borderRadius:11,background:"transparent",border:"1px solid rgba(255,255,255,0.08)",color:"#444",cursor:"pointer",fontSize:12}}>{lang==="de"?"Abbrechen":"Cancel"}</button>
          </div>
        </div>
      )}
    </div>
  );
}
