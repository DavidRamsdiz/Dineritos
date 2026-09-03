/* Dineritos — app completa. Generada; no editar a mano. */

/* ===== p3_engine.js ===== */
/* ==========================================================================
   Dineritos — finanzas personales
   Modelo portado del Excel "Dineritos Pro": mismo reparto proporcional,
   mismo ahorro con recuperacion, misma hucha. Todo se recalcula aqui.
   ========================================================================== */
(function(){
"use strict";

/* ---------------------------------------------------------------- constantes */
var MESES=["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO",
           "SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];
/* Orden de categorias tal y como esta en la hoja CONFIG del Excel */
var CATS_ORDEN=["PREVISTO","ANUAL","INQUISICIONES","CULTURA","VIDA DIARIA",
                "VIAJES","IMPREVISTO","AHORRO","HUCHA","GASTO HUCHA"];
/* Las que se pueden elegir al anotar un gasto (desplegable del Excel) */
var CATS_APUNTE=["VIDA DIARIA","INQUISICIONES","CULTURA","VIAJES","IMPREVISTO",
                 "HUCHA","GASTO HUCHA","AHORRO","INGRESO HUCHA"];
/* No son gasto del mes: el ahorro es patrimonio, el gasto hucha sale del fondo */
var NO_GASTO={"AHORRO":1,"GASTO HUCHA":1};
var FILAS_FIJAS={6:"COMUNES",7:"REC:0",8:"REC:1",9:"REC:2",10:"REC:3",11:"HACIENDA",
                 12:"COMPRAS",13:"VIDADIARIA",14:"AHORRO",15:"HUCHA",16:"BONO",18:"ALEMAN"};
var FUENTE_ESPERADA="Dineritos Pro.xlsx";

var S=null;   /* estado */
var V={tab:"mes",mes:null,anio:null,anioVista:null,dirty:false,tablaCat:false,
       desglose:null,desgloseAnio:null,anioModo:null};

/* ---------------------------------------------------------------- formato */
var nf2=new Intl.NumberFormat("es-ES",{minimumFractionDigits:2,maximumFractionDigits:2});
var nf0=new Intl.NumberFormat("es-ES",{maximumFractionDigits:0});
var nfp=new Intl.NumberFormat("es-ES",{maximumFractionDigits:1});
function n0(v){v=+v;return isFinite(v)?v:0;}
function menos(s){return String(s).replace(/^-/,"−");}
function eur(v){return menos(nf2.format(n0(v)))+" €";}
function eur0(v){return menos(nf0.format(Math.round(n0(v))))+" €";}
function eurS(v){v=n0(v);return (v>0?"+":v<0?"−":"")+eur(Math.abs(v));}
function eur0S(v){v=n0(v);return (v>0?"+":v<0?"−":"")+eur0(Math.abs(v));}
function pct(v){return nfp.format(n0(v)*100)+" %";}
function cap(m){return m?m.charAt(0)+m.slice(1).toLowerCase():"";}
function esc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;")
  .replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function mesLabel(mes,anio){return cap(mes)+" "+anio;}
function mkey(mes,anio){return mes+"-"+anio;}
function fechaCorta(iso){
  if(!iso)return"";
  var p=String(iso).split("-");
  return p.length===3?(+p[2])+" "+["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"][+p[1]-1]:String(iso);
}

/* ---------------------------------------------------------------- estado base */
function estadoVacio(){
  return {version:1,origen:null,savedAt:null,
    config:{sueldoDavid:0,sueldoCeline:0,pluxee:0,presupCompras:0,presupVidaDiaria:0,
            ahorroObjetivo:0,anioEnCurso:new Date().getFullYear(),fondoHucha:0,bonoAnual:0,
            mesBono:null,anioBono:null,cobeeAleman:0,cursoAleman:{},haciendaDefault:0},
    excepciones:[],reparto:[],gastosAnuales:[],recurrentes:[],mesesCurso:{},meses:[]};
}
function hayDatos(){return !!(S&&S.meses&&S.meses.length);}

/* ---------------------------------------------------------------- reparto */
var _cache={};
function invalidar(){_cache={};}

function repLineas(){
  if(_cache.rep)return _cache.rep;
  var ga=S.gastosAnuales||[];
  _cache.rep=(S.reparto||[]).map(function(l){
    var imp=n0(l.importe);
    if(l.linkAnual){
      var g=ga.filter(function(x){return x.concepto===l.linkAnual;})[0];
      if(g)imp=n0(g.importe)/12;
    }
    return {grupo:l.grupo,subconcepto:l.subconcepto,notas:l.notas,linkAnual:l.linkAnual,
            pagaDavidDirecto:!!l.pagaDavidDirecto,importe:n0(l.importe),imp:imp};
  });
  return _cache.rep;
}
function repTotales(){
  if(_cache.rept)return _cache.rept;
  var L=repLineas();
  var total=L.reduce(function(a,b){return a+b.imp;},0);
  var sd=n0(S.config.sueldoDavid),sc=n0(S.config.sueldoCeline),st=sd+sc;
  var pD=st?sd/st:0.5,pC=st?sc/st:0.5;
  var dir=L.filter(function(l){return l.pagaDavidDirecto;}).reduce(function(a,b){return a+b.imp;},0);
  var aD=total*pD,aC=total*pC;
  /* Agrupado por concepto, como la tabla G4:I12 del Excel */
  var grupos={},orden=[];
  L.forEach(function(l){
    if(!(l.grupo in grupos)){grupos[l.grupo]=0;orden.push(l.grupo);}
    grupos[l.grupo]+=l.imp;
  });
  _cache.rept={lineas:L,total:total,pD:pD,pC:pC,aD:aD,aC:aC,directo:dir,
    trD:aD-dir,trC:aC,mitad:total/2,difD:aD-total/2,
    grupos:orden.map(function(g){return {grupo:g,imp:grupos[g],pct:total?grupos[g]/total:0};})
      .sort(function(a,b){return b.imp-a.imp;})};
  return _cache.rept;
}

/* ---------------------------------------------------------------- mes: bases */
function excDe(mes,anio){
  return (S.excepciones||[]).filter(function(e){return e.mes===mes&&+e.anio===+anio;})[0]||null;
}
function bonoDe(mes,anio){
  var c=S.config;
  return (c.mesBono&&mes===c.mesBono&&+anio===+c.anioBono)?n0(c.bonoAnual):0;
}
function netoDe(mes,anio){
  var e=excDe(mes,anio);
  return (e?n0(e.neto):n0(S.config.sueldoDavid))+bonoDe(mes,anio);
}
function cursoDe(mes,anio){
  var lst=(S.mesesCurso||{})[String(anio)]||[];
  return lst.indexOf(mes)>=0?n0((S.config.cursoAleman||{})[String(anio)]):0;
}
function mesObj(mes,anio){
  return (S.meses||[]).filter(function(m){return m.mes===mes&&+m.anio===+anio;})[0]||null;
}
function mesesOrden(){
  return (S.meses||[]).slice().sort(function(a,b){
    return (a.anio-b.anio)||(MESES.indexOf(a.mes)-MESES.indexOf(b.mes));
  });
}
function anios(){
  var s={};(S.meses||[]).forEach(function(m){s[m.anio]=1;});
  return Object.keys(s).map(Number).sort(function(a,b){return a-b;});
}

/* Filas de plan del mes. sinAhorro=true evita la recursion al calcular el ahorro. */
function filasPlan(mes,anio,sinAhorro){
  var c=S.config,e=excDe(mes,anio),rt=repTotales(),rows=[];
  rows.push({id:"COMUNES",concepto:"Gastos comunes (mi aportación)",cat:"PREVISTO",plan:rt.aD,
    nota:"Reparto proporcional a los sueldos. Incluye el crédito del coche que pagas directamente."});
  (S.recurrentes||[]).forEach(function(r,i){
    rows.push({id:"REC:"+i,concepto:r.concepto,cat:r.categoria||"PREVISTO",plan:n0(r.importe),rec:i});
  });
  rows.push({id:"HACIENDA",concepto:"Hacienda",cat:"ANUAL",plan:e?n0(e.hacienda):n0(c.haciendaDefault),
    nota:e?"Provisión con la retención reducida de este mes":""});
  rows.push({id:"COMPRAS",concepto:"Compras (ropa, juegos, libros)",cat:"INQUISICIONES",plan:n0(c.presupCompras),
    nota:"Presupuesto: las compras reales se anotan abajo"});
  rows.push({id:"VIDADIARIA",concepto:"Vida diaria",cat:"VIDA DIARIA",plan:n0(c.presupVidaDiaria),
    nota:"Presupuesto: el gasto real se anota abajo"});
  if(!sinAhorro)rows.push({id:"AHORRO",concepto:"Ahorro",cat:"AHORRO",plan:ahorroPlan(mes,anio),
    nota:"Lo que cabe ahorrar este mes sin pasarte del objetivo anual"});
  rows.push({id:"HUCHA",concepto:"Fondo hucha (viajes + imprevistos)",cat:"HUCHA",plan:n0(c.fondoHucha),
    nota:"Se gasta durante el año en viajes e imprevistos; no cuenta como patrimonio"});
  var b=bonoDe(mes,anio);
  if(b>0)rows.push({id:"BONO",concepto:"Bono anual → inversión",cat:"AHORRO",plan:b,
    nota:"Ingreso extra que va directo a inversión; no toca el disponible del mes"});
  var al=cursoDe(mes,anio);
  if(al>0)rows.push({id:"ALEMAN",concepto:"Curso de alemán (Goethe-Institut)",cat:"CULTURA",plan:al});
  (S.gastosAnuales||[]).forEach(function(g,i){
    if(g.ambito==="PERSONAL"&&g.mes===mes)
      rows.push({id:"GA:"+i,concepto:g.concepto,cat:"ANUAL",plan:n0(g.importe),
        nota:"Gasto anual que cae en "+cap(mes),anual:true});
  });
  return rows;
}
function gastoPlanBase(mes,anio){
  return filasPlan(mes,anio,true).reduce(function(s,r){
    return NO_GASTO[r.cat]?s:s+r.plan;},0);
}

/* El ahorro del Excel: min(lo que sobra, objetivo pendiente / meses que faltan) */
function ahorroPlanMapa(){
  if(_cache.ap)return _cache.ap;
  var m={},porAnio={};
  mesesOrden().forEach(function(x){(porAnio[x.anio]=porAnio[x.anio]||[]).push(x);});
  Object.keys(porAnio).sort().forEach(function(y){
    var lst=porAnio[y],N=lst.length,acum=0,obj=n0(S.config.ahorroObjetivo);
    lst.forEach(function(x,i){
      var disp=(netoDe(x.mes,x.anio)-bonoDe(x.mes,x.anio))-gastoPlanBase(x.mes,x.anio);
      var techo=(obj*N-acum)/(N-i);
      var p=Math.max(0,Math.min(disp,techo));
      m[mkey(x.mes,x.anio)]=p;acum+=p;
    });
  });
  _cache.ap=m;return m;
}
function ahorroPlan(mes,anio){
  var v=ahorroPlanMapa()[mkey(mes,anio)];
  return v===undefined?0:v;
}

/* Valor real por defecto de una fila de plan cuando no hay nada anotado */
function realPorDefecto(id){
  return (id==="COMPRAS"||id==="VIDADIARIA")?null:"auto";
}
/* Importe real de una fila de plan: "auto" sigue al plan, un numero es un importe a mano */
function realDeFila(r,reales){
  var v=reales.hasOwnProperty(r.id)?reales[r.id]:realPorDefecto(r.id);
  var auto=(v==="auto"),real;
  if(auto)real=r.plan;
  else if(typeof v==="number")real=v;
  else real=null;
  return {auto:auto,real:real};
}

/* ---------------------------------------------------------------- mes: calculo */
function calcMes(mes,anio){
  var m=mesObj(mes,anio),reales=(m&&m.reales)||{},apuntes=(m&&m.apuntes)||[];
  var filas=filasPlan(mes,anio).map(function(r){
    var rr=realDeFila(r,reales),auto=rr.auto,real=rr.real;
    return {id:r.id,concepto:r.concepto,cat:r.cat,plan:r.plan,nota:r.nota||"",
      anual:!!r.anual,rec:r.rec,auto:auto,real:real,
      efec:real===null?0:real,dev:real===null?null:real-r.plan};
  });
  var gastoPlan=0,gastoReal=0,ahoPlan=0,ahoReal=0;
  filas.forEach(function(f){
    if(f.cat==="AHORRO"){ahoPlan+=f.plan;ahoReal+=f.efec;}
    else if(f.cat==="GASTO HUCHA"){/* sale del fondo, no del mes */}
    else{gastoPlan+=f.plan;gastoReal+=f.efec;}
  });
  apuntes.forEach(function(a){
    var v=n0(a.importe);
    if(a.categoria==="AHORRO")ahoReal+=v;
    else if(a.categoria==="GASTO HUCHA"){/* idem */}
    else gastoReal+=v;
  });
  var neto=netoDe(mes,anio),bono=bonoDe(mes,anio);
  /* Reparto por categoria: previsto y real */
  var porCat={};
  CATS_ORDEN.concat(["INGRESO HUCHA"]).forEach(function(c){porCat[c]={plan:0,real:0};});
  function acu(c,plan,real){
    if(!porCat[c])porCat[c]={plan:0,real:0};
    porCat[c].plan+=plan;porCat[c].real+=real;
  }
  filas.forEach(function(f){acu(f.cat,f.plan,f.efec);});
  apuntes.forEach(function(a){acu(a.categoria||"IMPREVISTO",0,n0(a.importe));});
  /* Hucha: la asignacion y lo gastado del fondo (virtual) y el ajuste fisico del mes.
     La hucha es una cuenta que no baja de cero: lo que no cubre lo adelantas tu y es
     gasto del mes; cuando lo recuperas, esa parte de la asignacion no sale de tu cuenta. */
  var h=huchaMes(mes,anio);
  var huchaIn=h.ent,huchaOut=h.sal;
  gastoReal+=h.adelanto-h.recuperado;
  /* Control de presupuestos */
  function sumApuntes(cat){
    return apuntes.reduce(function(s,a){return a.categoria===cat?s+n0(a.importe):s;},0);
  }
  var fCompras=filas.filter(function(f){return f.id==="COMPRAS";})[0];
  var fVida=filas.filter(function(f){return f.id==="VIDADIARIA";})[0];
  var presup=[
    {k:"compras",nom:"Compras (Inquisiciones)",pres:fCompras?fCompras.plan:0,gasto:sumApuntes("INQUISICIONES"),cats:["INQUISICIONES"]},
    {k:"vida",nom:"Vida diaria",pres:fVida?fVida.plan:0,gasto:sumApuntes("VIDA DIARIA"),cats:["VIDA DIARIA"]},
    {k:"ahorro",nom:"Ahorro del mes",pres:ahoPlan,gasto:ahoReal,bueno:true,cats:["AHORRO"]},
    {k:"hucha",nom:"Fondo hucha",pres:n0(S.config.fondoHucha),gasto:huchaIn,hucha:true,cats:["HUCHA","INGRESO HUCHA"]}
  ];
  return {mes:mes,anio:anio,key:mkey(mes,anio),filas:filas,apuntes:apuntes,
    neto:neto,bono:bono,pluxee:n0(S.config.pluxee),
    gastoPlan:gastoPlan,gastoReal:gastoReal,ahoPlan:ahoPlan,ahoReal:ahoReal,
    salidasPlan:gastoPlan+ahoPlan,salidasReal:gastoReal+ahoReal,
    balancePlan:neto-gastoPlan-ahoPlan,balanceReal:neto-gastoReal-ahoReal,
    porCat:porCat,huchaIn:huchaIn,huchaOut:huchaOut,hucha:h,presup:presup,
    anotado:apuntes.length>0};
}

/* Movimientos brutos de la hucha en un mes: la asignacion (fila HUCHA + apuntes
   INGRESO HUCHA) y lo gastado del fondo (apuntes GASTO HUCHA). Sin pasar por calcMes,
   porque calcMes necesita la serie para el ajuste fisico. */
function huchaFlujos(mes,anio){
  var m=mesObj(mes,anio),reales=(m&&m.reales)||{},apuntes=(m&&m.apuntes)||[];
  var fila=filasPlan(mes,anio).filter(function(f){return f.id==="HUCHA";})[0];
  var ent=fila?n0(realDeFila(fila,reales).real):0,sal=0;
  apuntes.forEach(function(a){
    if(a.categoria==="INGRESO HUCHA")ent+=n0(a.importe);
    if(a.categoria==="GASTO HUCHA")sal+=n0(a.importe);
  });
  return {ent:ent,sal:sal};
}

/* Serie de la hucha, mes a mes y en orden.
   La hucha es una cuenta de ahorro: NO puede bajar de cero. Se lleva el saldo "virtual"
   (asignado - gastado, como el Excel) y de ahi sale la realidad fisica:
     - deuda      = lo que la hucha te debe (el virtual en negativo)
     - adelanto   = lo que pones de tu bolsillo este mes porque la hucha se quedo a cero
     - recuperado = lo que te vuelve este mes: esa parte de la asignacion no entra fisicamente
     - entFis / salFis = lo que de verdad entra y sale de la cuenta
     - saldo      = saldo fisico, siempre >= 0
   Ejemplo: asignas 250 y gastas 300 -> saldo 0, adelantas 50. Al mes siguiente asignas 250:
   50 vuelven a ti, entran 200 y la hucha queda en 200. */
function serieHucha(){
  if(_cache.hu)return _cache.hu;
  var virtual=0,saldo=0,deuda=0;
  _cache.hu=mesesOrden().map(function(m){
    var f=huchaFlujos(m.mes,m.anio);
    virtual+=f.ent-f.sal;
    var deudaNueva=Math.max(0,-virtual);
    var adelanto=Math.max(0,deudaNueva-deuda),recuperado=Math.max(0,deuda-deudaNueva);
    deuda=deudaNueva;
    var saldoIni=saldo;
    saldo=Math.max(0,virtual);
    return {mes:m.mes,anio:m.anio,key:mkey(m.mes,m.anio),ent:f.ent,sal:f.sal,
      entFis:f.ent-recuperado,salFis:f.sal-adelanto,adelanto:adelanto,recuperado:recuperado,
      deuda:deuda,saldoIni:saldoIni,saldo:saldo,virtual:virtual};
  });
  return _cache.hu;
}
var HUCHA_VACIA={ent:0,sal:0,entFis:0,salFis:0,adelanto:0,recuperado:0,deuda:0,saldoIni:0,saldo:0,virtual:0};
function huchaMes(mes,anio){
  var k=mkey(mes,anio);
  return serieHucha().filter(function(h){return h.key===k;})[0]||HUCHA_VACIA;
}
function huchaSaldoTras(mes,anio){return huchaMes(mes,anio).saldo;}

/* Filas de todos los meses de un ano, en orden, con el acumulado de ahorro */
function filasAnio(anio){
  var ck="fa"+anio;if(_cache[ck])return _cache[ck];
  var hu=serieHucha(),acum=0;
  _cache[ck]=mesesOrden().filter(function(m){return +m.anio===+anio;}).map(function(m){
    var c=calcMes(m.mes,m.anio);
    acum+=c.ahoReal;
    var h=hu.filter(function(x){return x.key===c.key;})[0]||HUCHA_VACIA;
    return {mes:m.mes,anio:m.anio,key:c.key,neto:c.neto,gastoPlan:c.gastoPlan,gastoReal:c.gastoReal,
      ahoPlan:c.ahoPlan,ahoReal:c.ahoReal,balancePlan:c.balancePlan,balanceReal:c.balanceReal,
      acum:acum,porCat:c.porCat,huEnt:h.ent,huSal:h.sal,huSaldo:h.saldo,
      huAdel:h.adelanto,huRec:h.recuperado,huDeuda:h.deuda,anotado:c.anotado};
  });
  return _cache[ck];
}
function sumaFilas(filas){
  return filas.reduce(function(a,f){
    a.neto+=f.neto;a.gastoPlan+=f.gastoPlan;a.gastoReal+=f.gastoReal;
    a.ahoPlan+=f.ahoPlan;a.ahoReal+=f.ahoReal;
    a.balancePlan+=f.balancePlan;a.balanceReal+=f.balanceReal;
    a.huEnt+=f.huEnt;a.huSal+=f.huSal;a.huAdel+=f.huAdel;a.huRec+=f.huRec;
    a.bono+=bonoDe(f.mes,f.anio);a.meses++;return a;
  },{neto:0,gastoPlan:0,gastoReal:0,ahoPlan:0,ahoReal:0,balancePlan:0,balanceReal:0,
     huEnt:0,huSal:0,huAdel:0,huRec:0,bono:0,meses:0});
}

/* Resumen de un ano. Con hastaMes (p. ej. "SEPTIEMBRE") solo cuenta los meses hasta ese,
   incluido: es la vista "hasta hoy". Lo que queda del ano va en `resto`, segun el plan. */
function resumenAnio(anio,hastaMes){
  var ck="res"+anio+"|"+(hastaMes||"");if(_cache[ck])return _cache[ck];
  var todas=filasAnio(anio),filas=todas,resto=[];
  if(hastaMes){
    var hi=MESES.indexOf(hastaMes);
    filas=todas.filter(function(f){return MESES.indexOf(f.mes)<=hi;});
    resto=todas.slice(filas.length);
  }
  var tot=sumaFilas(filas);
  var cats=CATS_ORDEN.map(function(c){
    return {cat:c,
      real:filas.reduce(function(s,f){return s+(f.porCat[c]?f.porCat[c].real:0);},0),
      plan:filas.reduce(function(s,f){return s+(f.porCat[c]?f.porCat[c].plan:0);},0)};
  });
  var obj=n0(S.config.ahorroObjetivo);
  var hu=serieHucha(),saldoIni=0;
  var idx=hu.map(function(h){return h.key;}).indexOf(filas.length?filas[0].key:"");
  if(idx>0)saldoIni=hu[idx-1].saldo;
  var out={anio:anio,filas:filas,tot:tot,cats:cats,objetivo:obj*filas.length,
    objetivoAnual:obj*todas.length,mesesAnio:todas.length,
    hasta:hastaMes?(filas.length?filas[filas.length-1].mes:null):null,
    resto:sumaFilas(resto),
    conseguidoPlan:filas.reduce(function(s,f){return s+f.ahoPlan;},0),
    conseguidoReal:tot.ahoReal,saldoIni:saldoIni,
    saldoFin:filas.length?filas[filas.length-1].huSaldo:saldoIni,
    deudaFin:filas.length?filas[filas.length-1].huDeuda:0,
    bono:tot.bono,
    meses:filas.length};
  _cache[ck]=out;return out;
}

/* Desglose de una o varias categorias: los apuntes (y las filas del plan con importe) que
   las componen. Con `mes` solo ese mes; con `hastaMes` solo hasta ese mes del ano. */
function desgloseCat(cats,anio,mes,hastaMes){
  var out=[],hi=hastaMes?MESES.indexOf(hastaMes):99;
  mesesOrden().forEach(function(m){
    if(+m.anio!==+anio||(mes&&m.mes!==mes)||MESES.indexOf(m.mes)>hi)return;
    var c=calcMes(m.mes,m.anio),pref=mes?"":cap(m.mes)+" · ";
    c.filas.forEach(function(f){
      if(cats.indexOf(f.cat)>=0&&Math.abs(f.efec)>0.005)
        out.push({concepto:f.concepto,importe:f.efec,sub:pref+"del plan del mes",cat:f.cat});
    });
    c.apuntes.forEach(function(a){
      if(cats.indexOf(a.categoria)>=0&&Math.abs(n0(a.importe))>0.005)
        out.push({concepto:a.concepto,importe:n0(a.importe),
          sub:a.fecha?pref+fechaCorta(a.fecha):(mes?"":cap(m.mes)),cat:a.categoria});
    });
  });
  return out;
}

/* ==========================================================================
   Lector de .xlsx en la propia pagina (ZIP + XML, sin librerias externas)
   ========================================================================== */
function leerZip(ab){
  var dv=new DataView(ab),u8=new Uint8Array(ab),eo=-1;
  var min=Math.max(0,u8.length-66000);
  for(var i=u8.length-22;i>=min;i--){if(dv.getUint32(i,true)===0x06054b50){eo=i;break;}}
  if(eo<0)throw new Error("El archivo no parece un .xlsx válido.");
  var n=dv.getUint16(eo+10,true),p=dv.getUint32(eo+16,true),out={};
  for(var k=0;k<n;k++){
    if(dv.getUint32(p,true)!==0x02014b50)throw new Error("El índice interno del .xlsx está dañado.");
    var metodo=dv.getUint16(p+10,true),crc=dv.getUint32(p+16,true);
    var csize=dv.getUint32(p+20,true),usize=dv.getUint32(p+24,true);
    var nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true);
    var lho=dv.getUint32(p+42,true);
    var nombre=new TextDecoder("utf-8").decode(u8.subarray(p+46,p+46+nl));
    var lnl=dv.getUint16(lho+26,true),lel=dv.getUint16(lho+28,true);
    var ini=lho+30+lnl+lel;
    /* crc y usize se guardan para poder recopiar la entrada verbatim al reescribir
       el libro: asi las partes binarias (imagenes de los dibujos) no se tocan. */
    out[nombre]={metodo:metodo,raw:u8.slice(ini,ini+csize),crc:crc,usize:usize};
    p+=46+nl+el+cl;
  }
  return out;
}
function inflar(e){
  if(!e)return Promise.resolve(null);
  if(e.metodo===0)return Promise.resolve(new TextDecoder("utf-8").decode(e.raw));
  if(typeof DecompressionStream==="undefined")
    return Promise.reject(new Error("Este navegador no puede abrir .xlsx (le falta DecompressionStream). Prueba con Chrome, Edge o Safari actualizados."));
  var s=new Blob([e.raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Response(s).text();
}
function xml(txt){
  if(!txt)return null;
  var d=new DOMParser().parseFromString(txt,"application/xml");
  if(d.getElementsByTagName("parsererror").length)throw new Error("No he podido leer el XML interno del .xlsx.");
  return d;
}
function celdas(doc,shared){
  var map={},cs=doc.getElementsByTagName("c");
  for(var i=0;i<cs.length;i++){
    var c=cs[i],ref=c.getAttribute("r");if(!ref)continue;
    var t=c.getAttribute("t")||"n",f=null,v=null;
    for(var j=0;j<c.childNodes.length;j++){
      var ch=c.childNodes[j];if(ch.nodeType!==1)continue;
      var tn=ch.nodeName.replace(/^.*:/,"");
      if(tn==="f")f=ch.textContent||"";
      else if(tn==="v")v=ch.textContent;
      else if(tn==="is"){
        var ts=ch.getElementsByTagName("t"),s="";
        for(var q=0;q<ts.length;q++)s+=ts[q].textContent;
        v=s;
      }
    }
    var val;
    if(v===null)val=null;
    else if(t==="s"){var ix=parseInt(v,10);val=shared[ix]===undefined?null:shared[ix];}
    else if(t==="inlineStr"||t==="str")val=v;
    else if(t==="b")val=(v==="1");
    else if(t==="e")val=null;
    else{var nv=parseFloat(v);val=isNaN(nv)?v:nv;}
    map[ref]={v:val,f:f};
  }
  return map;
}
function gv(c,r){var x=c[r];return x?x.v:null;}
function gn(c,r,d){var v=gv(c,r);return typeof v==="number"?v:(d===undefined?0:d);}
function gs(c,r){var v=gv(c,r);return v==null?"":String(v).trim();}
function esFormula(c,r){var x=c[r];return !!(x&&x.f!==null&&x.f!==undefined);}
function formula(c,r){var x=c[r];return x&&x.f?x.f:"";}
function serieAFecha(n){
  if(typeof n==="string"){
    var m=/^(\d{4})-(\d{2})-(\d{2})/.exec(n.trim());
    return m?m[0]:null;
  }
  if(typeof n!=="number"||n<20000||n>90000)return null;
  var ms=Date.UTC(1899,11,30)+Math.round(n)*86400000;
  return new Date(ms).toISOString().slice(0,10);
}
/* Si E2 no trae valor cacheado, el ano se saca del nombre de la hoja */
function anioDeNombre(nom){
  var m=/(\d{4})\s*$/.exec(nom);
  if(m)return +m[1];
  m=/(\d{2})\s*$/.exec(nom);
  return m?2000+ +m[1]:null;
}

function importarLibro(hojas){
  var cfg=hojas["CONFIG"];
  if(!cfg)throw new Error("No encuentro la hoja CONFIG. ¿Es el Excel “Dineritos Pro”?");
  var st=estadoVacio();
  st.origen=FUENTE_ESPERADA;

  /* --- CONFIG: parametros --- */
  var c=st.config;
  c.sueldoDavid=gn(cfg,"F5");c.sueldoCeline=gn(cfg,"F6");c.pluxee=gn(cfg,"F7");
  c.presupCompras=gn(cfg,"F8");c.presupVidaDiaria=gn(cfg,"F9");c.ahorroObjetivo=gn(cfg,"F10");
  c.anioEnCurso=gn(cfg,"F11",new Date().getFullYear());c.fondoHucha=gn(cfg,"F12");
  c.bonoAnual=gn(cfg,"F13");c.mesBono=gs(cfg,"F14")||null;c.cobeeAleman=gn(cfg,"F15");
  c.haciendaDefault=165;

  /* --- CONFIG: excepciones de neto y provision de Hacienda --- */
  for(var r=18;r<=60;r++){
    var mm=gs(cfg,"G"+r).toUpperCase();
    if(MESES.indexOf(mm)<0)continue;
    var an=gn(cfg,"H"+r,0);if(!an)continue;
    st.excepciones.push({mes:mm,anio:an,neto:gn(cfg,"I"+r),hacienda:gn(cfg,"J"+r),_fila:r});
  }

  /* --- GASTOS_ANUALES --- */
  var ga=hojas["GASTOS_ANUALES"];
  if(ga){
    for(var r2=4;r2<=60;r2++){
      var co=gs(ga,"B"+r2);
      if(!co||/^TOTAL/i.test(co))continue;
      st.gastosAnuales.push({concepto:co,importe:gn(ga,"C"+r2),
        mes:gs(ga,"D"+r2).toUpperCase(),ambito:gs(ga,"E"+r2).toUpperCase(),_fila:r2});
    }
  }

  /* --- REPARTO: gastos comunes --- */
  var rp=hojas["REPARTO"];
  if(rp){
    /* La tabla empieza en la fila 5 y termina en la fila del TOTAL: a partir de ahi
       vienen los bloques de reparto, cuya columna D tambien lleva importes. */
    for(var r3=5;r3<=60;r3++){
      var g=gs(rp,"B"+r3);
      if(/^\s*=?\s*TOTAL/i.test(g))break;
      if(!g)continue;
      if(typeof gv(rp,"D"+r3)!=="number")continue;
      var sub=gs(rp,"C"+r3),notas=gs(rp,"E"+r3);
      var link=null,fx=formula(rp,"D"+r3),mfx=/GASTOS_ANUALES!\$?F\$?(\d+)/i.exec(fx);
      if(mfx&&ga)link=gs(ga,"B"+mfx[1])||null;
      st.reparto.push({grupo:g,subconcepto:sub,importe:gn(rp,"D"+r3),notas:notas,linkAnual:link,
        pagaDavidDirecto:/paga\s+david/i.test(notas)||/cr[eé]dito/i.test(sub),_fila:r3});
    }
  }

  /* --- Hojas mensuales: se reconocen porque B2 dice "MES" --- */
  var nombres=Object.keys(hojas),aniosVistos={};
  nombres.forEach(function(nom){
    var h=hojas[nom];
    if(gs(h,"B2").toUpperCase()!=="MES")return;
    var mes=gs(h,"C2").toUpperCase();
    if(MESES.indexOf(mes)<0)return;
    var anio=gn(h,"E2",0)||anioDeNombre(nom)||c.anioEnCurso;
    aniosVistos[anio]=1;
    /* recurrentes: se toman de la primera hoja mensual que aparezca */
    if(!st.recurrentes.length){
      [7,8,9,10].forEach(function(rr){
        var nm=gs(h,"B"+rr);
        if(nm)st.recurrentes.push({concepto:nm,categoria:gs(h,"C"+rr)||"PREVISTO",importe:gn(h,"D"+rr)});
      });
    }
    /* importes reales de las filas fijas */
    var reales={};
    Object.keys(FILAS_FIJAS).forEach(function(rr){
      var id=FILAS_FIJAS[rr],ref="E"+rr;
      if(esFormula(h,ref))reales[id]="auto";
      else{var v=gv(h,ref);reales[id]=(typeof v==="number")?v:null;}
    });
    /* gastos anuales personales de este mes: filas 39-43, en el orden de la hoja */
    var idxGA=[];
    st.gastosAnuales.forEach(function(x,i){if(x.ambito==="PERSONAL"&&x.mes===mes)idxGA.push(i);});
    idxGA.forEach(function(gi,k){
      var ref2="E"+(39+k);
      if(esFormula(h,ref2))reales["GA:"+gi]="auto";
      else{var v2=gv(h,ref2);reales["GA:"+gi]=(typeof v2==="number")?v2:null;}
    });
    /* apuntes libres: filas 19-37 */
    var apuntes=[];
    for(var rr2=19;rr2<=37;rr2++){
      var cn=gs(h,"B"+rr2);if(!cn)continue;
      apuntes.push({concepto:cn,categoria:gs(h,"C"+rr2)||"IMPREVISTO",importe:gn(h,"E"+rr2),
        fecha:serieAFecha(gv(h,"G"+rr2)),notas:gs(h,"H"+rr2)});
    }
    /* meses con curso de aleman: los que traen importe en D18 */
    if(gn(h,"D18")>0||/CONFIG!\$?F\$?(16|18)/i.test(formula(h,"D18"))){
      var ky=String(anio);
      (st.mesesCurso[ky]=st.mesesCurso[ky]||[]).push(mes);
    }
    st.meses.push({mes:mes,anio:anio,reales:reales,apuntes:apuntes,hojaOrigen:nom});
  });
  if(!st.meses.length)throw new Error("No he encontrado ninguna hoja mensual (las reconozco porque la celda B2 dice “MES”).");

  /* importe del curso de aleman por ano: F16 el primero, F18 el segundo */
  var ys=Object.keys(aniosVistos).map(Number).sort(function(a,b){return a-b;});
  if(ys[0])st.config.cursoAleman[String(ys[0])]=gn(cfg,"F16");
  if(ys[1])st.config.cursoAleman[String(ys[1])]=gn(cfg,"F18");

  /* Hacienda por defecto: el importe mas repetido entre las excepciones (ignorando los ceros) */
  var cuenta={},mejor=null;
  st.excepciones.forEach(function(x){
    if(!x.hacienda)return;
    cuenta[x.hacienda]=(cuenta[x.hacienda]||0)+1;
    if(mejor===null||cuenta[x.hacienda]>cuenta[mejor])mejor=x.hacienda;
  });
  if(mejor!==null)st.config.haciendaDefault=mejor;

  /* mes del bono: el ano es el de la ultima excepcion que lo contenga */
  if(st.config.mesBono){
    var cand=st.excepciones.filter(function(e){return e.mes===st.config.mesBono;});
    st.config.anioBono=cand.length?cand[cand.length-1].anio:(ys[ys.length-1]||null);
  }
  st.version=1;
  return st;
}

function abrirXlsx(file){
  return file.arrayBuffer().then(function(ab){
    var z=leerZip(ab);
    function pick(re){
      var k=Object.keys(z).filter(function(n){return re.test(n);});
      return k.length?z[k[0]]:null;
    }
    return Promise.all([
      inflar(pick(/^xl\/workbook\.xml$/)),
      inflar(pick(/^xl\/_rels\/workbook\.xml\.rels$/)),
      inflar(pick(/^xl\/sharedStrings\.xml$/))
    ]).then(function(res){
      var wb=xml(res[0]),rels=xml(res[1]),sst=xml(res[2]);
      if(!wb)throw new Error("El .xlsx no tiene xl/workbook.xml; puede estar corrupto.");
      var shared=[];
      if(sst){
        var sis=sst.getElementsByTagName("si");
        for(var i=0;i<sis.length;i++){
          var ts=sis[i].getElementsByTagName("t"),s="";
          for(var j=0;j<ts.length;j++)s+=ts[j].textContent;
          shared.push(s);
        }
      }
      var relMap={};
      if(rels){
        var rl=rels.getElementsByTagName("Relationship");
        for(var k2=0;k2<rl.length;k2++)relMap[rl[k2].getAttribute("Id")]=rl[k2].getAttribute("Target");
      }
      var NSR="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      var sh=wb.getElementsByTagName("sheet"),tareas=[],nombres=[];
      for(var s2=0;s2<sh.length;s2++){
        var nom=sh[s2].getAttribute("name");
        var rid=sh[s2].getAttributeNS(NSR,"id")||sh[s2].getAttribute("r:id");
        var tgt=relMap[rid]||("worksheets/sheet"+(s2+1)+".xml");
        tgt=String(tgt).replace(/^\/?xl\//,"").replace(/^\//,"");
        var ent=z["xl/"+tgt]||z[tgt];
        nombres.push(nom);tareas.push(inflar(ent));
      }
      return Promise.all(tareas).then(function(txts){
        var hojas={};
        txts.forEach(function(t,i){
          if(!t)return;
          try{hojas[nombres[i]]=celdas(xml(t),shared);}catch(e){}
        });
        return importarLibro(hojas);
      });
    });
  });
}

/* ==========================================================================
   Exportaciones
   ========================================================================== */
function csvEsc(v){
  var s=v==null?"":String(v);
  return /[";\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;
}
function csv(filas){
  return "﻿"+filas.map(function(f){return f.map(csvEsc).join(";");}).join("\r\n");
}
function numES(v){return String(n0(v).toFixed(2)).replace(".",",");}

function csvApuntes(){
  var out=[["ANO","MES","FECHA","CONCEPTO","CATEGORIA","IMPORTE","NOTAS"]];
  mesesOrden().forEach(function(m){
    (m.apuntes||[]).forEach(function(a){
      out.push([m.anio,cap(m.mes),a.fecha||"",a.concepto,a.categoria,numES(a.importe),a.notas||""]);
    });
  });
  return csv(out);
}
function csvResumen(anio){
  var R=resumenAnio(anio);
  var out=[["MES","INGRESOS","GASTOS PREV.","GASTOS REAL","AHORRO","BALANCE PREV.","BALANCE REAL",
            "AHORRO ACUM.","HUCHA ENTRA","HUCHA SALE","HUCHA TU BOLSILLO","HUCHA SALDO"]];
  R.filas.forEach(function(f){
    out.push([cap(f.mes),numES(f.neto),numES(f.gastoPlan),numES(f.gastoReal),numES(f.ahoReal),
      numES(f.balancePlan),numES(f.balanceReal),numES(f.acum),numES(f.huEnt),numES(f.huSal),
      numES(f.huRec-f.huAdel),numES(f.huSaldo)]);
  });
  out.push(["TOTAL "+anio,numES(R.tot.neto),numES(R.tot.gastoPlan),numES(R.tot.gastoReal),
    numES(R.tot.ahoReal),numES(R.tot.balancePlan),numES(R.tot.balanceReal),
    numES(R.tot.ahoReal),numES(R.tot.huEnt),numES(R.tot.huSal),numES(R.tot.huRec-R.tot.huAdel),numES(R.saldoFin)]);
  out.push([]);
  out.push(["GASTO POR CATEGORIA","REAL","PREVISTO"]);
  R.cats.forEach(function(c){out.push([c.cat,numES(c.real),numES(c.plan)]);});
  return csv(out);
}
/* Filas listas para pegar en la zona de apuntes (B19) de la hoja del mes */
function tsvMes(mes,anio){
  var m=mesObj(mes,anio);if(!m)return "";
  return (m.apuntes||[]).map(function(a){
    return [a.concepto,a.categoria,"",numES(a.importe)].join("\t");
  }).join("\n");
}

/* Exportar al ambito global de la app */
window.DIN={
  MESES:MESES,CATS_ORDEN:CATS_ORDEN,CATS_APUNTE:CATS_APUNTE,NO_GASTO:NO_GASTO,
  n0:n0,eur:eur,eur0:eur0,eurS:eurS,eur0S:eur0S,pct:pct,cap:cap,esc:esc,
  mesLabel:mesLabel,mkey:mkey,fechaCorta:fechaCorta,numES:numES,
  estadoVacio:estadoVacio,invalidar:invalidar,
  repLineas:repLineas,repTotales:repTotales,
  excDe:excDe,bonoDe:bonoDe,netoDe:netoDe,cursoDe:cursoDe,mesObj:mesObj,
  mesesOrden:mesesOrden,anios:anios,filasPlan:filasPlan,ahorroPlan:ahorroPlan,
  calcMes:calcMes,serieHucha:serieHucha,huchaMes:huchaMes,huchaSaldoTras:huchaSaldoTras,resumenAnio:resumenAnio,desgloseCat:desgloseCat,
  abrirXlsx:abrirXlsx,csvApuntes:csvApuntes,csvResumen:csvResumen,tsvMes:tsvMes,
  leerZip:leerZip,inflar:inflar,
  hayDatos:hayDatos,V:V,
  FILAS_FIJAS:FILAS_FIJAS,AP_INI:19,AP_FIN:37,GA_INI:39,GA_FIN:43,
  get S(){return S;},set S(x){S=x;invalidar();}
};
})();


/* ===== p8_escritura.js ===== */
/* ==========================================================================
   Escritura en el .xlsx desde el navegador.

   La API de Excel de Graph (/workbook/...) NO sirve con OneDrive personal:
   Microsoft solo la soporta en OneDrive de empresa. Asi que se hace lo mismo
   que el script de Python: el .xlsx es un ZIP de XML, se cambian SOLO las
   celdas que toca, se vuelve a empaquetar y se sube.

   Todo lo que no se toca se copia VERBATIM (los mismos bytes ya comprimidos),
   asi que graficos, dibujos, imagenes, formato condicional y validaciones
   quedan exactamente igual.
   ========================================================================== */
(function(){
"use strict";
var D=window.DIN;

/* ---------------------------------------------------------------- CRC32 */
var TCRC=(function(){
  var t=new Uint32Array(256);
  for(var n=0;n<256;n++){
    var c=n;
    for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);
    t[n]=c>>>0;
  }
  return t;
})();
function crc32(u8){
  var c=0xFFFFFFFF;
  for(var i=0;i<u8.length;i++)c=TCRC[(c^u8[i])&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0;
}

/* ---------------------------------------------------------------- ZIP */
function comprimir(u8){
  if(typeof CompressionStream==="undefined")return Promise.resolve(null);
  try{
    var s=new Blob([u8]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return new Response(s).arrayBuffer().then(
      function(b){return new Uint8Array(b);},function(){return null;});
  }catch(x){return Promise.resolve(null);}
}
var DOS_HORA=0,DOS_FECHA=((2020-1980)<<9)|(1<<5)|1;

/* entradas: {nombre: {crudo,metodo,crc,usize} | {bytes:Uint8Array}}
   Las que llegan con `crudo` se copian tal cual; las que llegan con `bytes`
   se comprimen aqui. */
function escribirZip(orden,entradas){
  var tareas=orden.map(function(nombre){
    var e=entradas[nombre];
    if(e.crudo)return Promise.resolve({nombre:nombre,metodo:e.metodo,crc:e.crc,
      usize:e.usize,cuerpo:e.crudo});
    var datos=e.bytes;
    return comprimir(datos).then(function(comp){
      var usar=comp&&comp.length<datos.length;
      return {nombre:nombre,metodo:usar?8:0,crc:crc32(datos),usize:datos.length,
              cuerpo:usar?comp:datos};
    });
  });
  return Promise.all(tareas).then(function(ents){
    var enc=new TextEncoder(),trozos=[],central=[],offset=0;
    ents.forEach(function(e){
      var nom=enc.encode(e.nombre);
      var lh=new Uint8Array(30+nom.length),dv=new DataView(lh.buffer);
      dv.setUint32(0,0x04034b50,true);
      dv.setUint16(4,20,true);dv.setUint16(6,0,true);
      dv.setUint16(8,e.metodo,true);
      dv.setUint16(10,DOS_HORA,true);dv.setUint16(12,DOS_FECHA,true);
      dv.setUint32(14,e.crc,true);
      dv.setUint32(18,e.cuerpo.length,true);
      dv.setUint32(22,e.usize,true);
      dv.setUint16(26,nom.length,true);dv.setUint16(28,0,true);
      lh.set(nom,30);
      trozos.push(lh,e.cuerpo);

      var cd=new Uint8Array(46+nom.length),dc=new DataView(cd.buffer);
      dc.setUint32(0,0x02014b50,true);
      dc.setUint16(4,20,true);dc.setUint16(6,20,true);dc.setUint16(8,0,true);
      dc.setUint16(10,e.metodo,true);
      dc.setUint16(12,DOS_HORA,true);dc.setUint16(14,DOS_FECHA,true);
      dc.setUint32(16,e.crc,true);
      dc.setUint32(20,e.cuerpo.length,true);
      dc.setUint32(24,e.usize,true);
      dc.setUint16(28,nom.length,true);
      dc.setUint16(30,0,true);dc.setUint16(32,0,true);
      dc.setUint16(34,0,true);dc.setUint16(36,0,true);
      dc.setUint32(38,0,true);
      dc.setUint32(42,offset,true);
      cd.set(nom,46);
      central.push(cd);
      offset+=lh.length+e.cuerpo.length;
    });
    var tamCentral=central.reduce(function(a,b){return a+b.length;},0);
    var eo=new Uint8Array(22),de=new DataView(eo.buffer);
    de.setUint32(0,0x06054b50,true);
    de.setUint16(4,0,true);de.setUint16(6,0,true);
    de.setUint16(8,ents.length,true);de.setUint16(10,ents.length,true);
    de.setUint32(12,tamCentral,true);
    de.setUint32(16,offset,true);
    de.setUint16(20,0,true);
    var todo=trozos.concat(central,[eo]);
    var largo=todo.reduce(function(a,b){return a+b.length;},0);
    var salida=new Uint8Array(largo),p=0;
    todo.forEach(function(t){salida.set(t,p);p+=t.length;});
    return salida;
  });
}

/* ---------------------------------------------------------------- celdas */
function xesc(s){
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}
function colNum(s){
  var n=0;
  for(var i=0;i<s.length;i++)n=n*26+(s.charCodeAt(i)-64);
  return n;
}
function colStr(n){
  var s="";
  while(n>0){var r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}
  return s;
}
function celdasDeRango(dir){
  var m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(dir);
  if(!m)return null;
  var c1=colNum(m[1]),f1=+m[2];
  var c2=m[3]?colNum(m[3]):c1,f2=m[4]?+m[4]:f1;
  var filas=[];
  for(var f=f1;f<=f2;f++){
    var fila=[];
    for(var c=c1;c<=c2;c++)fila.push(colStr(c)+f);
    filas.push(fila);
  }
  return filas;
}
function localizar(xml,ref){
  var m=new RegExp('<c r="'+ref+'"(?![0-9])[^>]*/>').exec(xml),txt;
  if(m)txt=m[0];
  else{
    m=new RegExp('<c r="'+ref+'"(?![0-9])[^>]*>[\\s\\S]*?</c>').exec(xml);
    if(!m)return null;
    txt=m[0];
  }
  var est=/\ss="(\d+)"/.exec(txt);
  return {ini:m.index,fin:m.index+txt.length,txt:txt,estilo:est?est[1]:null};
}
function mismaFormula(txt,valor){
  var m=/<f[^>]*>([\s\S]*?)<\/f>/.exec(txt);
  if(!m)return false;
  var a=m[1].replace(/\$/g,"").replace(/&amp;/g,"&").toUpperCase();
  var b=valor.slice(1).replace(/\$/g,"").toUpperCase();
  return a===b;
}
function nuevaCelda(ref,estilo,valor){
  var s=estilo?' s="'+estilo+'"':"";
  if(typeof valor==="string"&&valor.charAt(0)==="=")
    return '<c r="'+ref+'"'+s+'><f>'+xesc(valor.slice(1))+'</f></c>';
  if(valor===null||valor===undefined||valor==="")
    return '<c r="'+ref+'"'+s+'/>';
  if(typeof valor==="number"&&isFinite(valor))
    return '<c r="'+ref+'"'+s+'><v>'+(Math.round(valor*1e10)/1e10)+'</v></c>';
  return '<c r="'+ref+'"'+s+' t="inlineStr"><is><t>'+xesc(valor)+'</t></is></c>';
}

/* ---------------------------------------------------------------- libro */
function mapaHojas(entradas){
  return Promise.all([
    D.inflar(entradas["xl/workbook.xml"]),
    D.inflar(entradas["xl/_rels/workbook.xml.rels"])
  ]).then(function(r){
    var wb=r[0]||"",rels=r[1]||"",rmap={},mapa={},m;
    var re=/Id="(rId\d+)"[^>]*Target="([^"]+)"/g;
    while((m=re.exec(rels)))rmap[m[1]]=m[2];
    var rs=/<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"/g;
    while((m=rs.exec(wb))){
      var tgt=String(rmap[m[2]]||"").replace(/^\/?xl\//,"").replace(/^\//,"");
      var clave="xl/"+tgt;
      if(entradas[clave]){
        var nom=m[1].replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">")
                    .replace(/&quot;/g,'"').replace(/&apos;/g,"'");
        mapa[nom]=clave;
      }
    }
    return mapa;
  });
}
function forzarRecalculo(wb){
  if(wb.indexOf('fullCalcOnLoad="1"')>=0)return wb;
  if(/<calcPr[^>]*\/>/.test(wb))
    return wb.replace(/<calcPr([^>]*?)\/>/,'<calcPr$1 fullCalcOnLoad="1"/>');
  if(wb.indexOf("<calcPr")>=0)
    return wb.replace(/<calcPr([^>]*)>/,'<calcPr$1 fullCalcOnLoad="1">');
  return wb.replace("</workbook>",'<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
}

/* Aplica el plan sobre los bytes del libro y devuelve los bytes nuevos. */
function aplicar(ab,plan){
  var entradas,orden;
  try{
    entradas=D.leerZip(ab);
    orden=Object.keys(entradas);
  }catch(err){
    return Promise.reject(new Error("No he podido abrir el libro para escribir: "+err.message));
  }
  if(!entradas["xl/workbook.xml"])
    return Promise.reject(new Error("El libro no tiene xl/workbook.xml."));

  return mapaHojas(entradas).then(function(mapa){
    var hojasATocar={},faltan=[];
    plan.forEach(function(d){
      var clave=mapa[d.hoja];
      if(clave)hojasATocar[clave]=true;
      else if(faltan.indexOf(d.hoja)<0)faltan.push(d.hoja);
    });
    var claves=Object.keys(hojasATocar);
    /* workbook.xml se toca siempre, para marcar el recalculo al abrir */
    var aLeer=claves.concat(["xl/workbook.xml"]);
    return Promise.all(aLeer.map(function(k){return D.inflar(entradas[k]);}))
     .then(function(textos){
       var xmls={};
       aLeer.forEach(function(k,i){xmls[k]=textos[i];});
       var escritas=0,noEncontradas=[];
       plan.forEach(function(d){
         var clave=mapa[d.hoja];
         if(!clave)return;
         var rejilla=celdasDeRango(d.dir);
         if(!rejilla){noEncontradas.push(d.hoja+"!"+d.dir+" (rango inválido)");return;}
         var datos=d.cuerpo.values||d.cuerpo.formulas||[];
         var esF=!!d.cuerpo.formulas;
         var xml=xmls[clave];
         for(var f=0;f<rejilla.length;f++){
           for(var c=0;c<rejilla[f].length;c++){
             var ref=rejilla[f][c],v=(datos[f]||[])[c];
             if(v===undefined)continue;
             if(esF&&typeof v==="string"&&v!==""&&v.charAt(0)!=="=")v="="+v;
             var loc=localizar(xml,ref);
             if(!loc){noEncontradas.push(d.hoja+"!"+ref);continue;}
             /* Si la celda ya lleva esa misma formula (aunque escrita con $), se deja
                tal cual: asi el archivo cambia lo menos posible. */
             if(typeof v==="string"&&v.charAt(0)==="="&&mismaFormula(loc.txt,v))continue;
             var nueva=nuevaCelda(ref,loc.estilo,v);
             if(nueva===loc.txt)continue;
             xml=xml.slice(0,loc.ini)+nueva+xml.slice(loc.fin);
             escritas++;
           }
         }
         xmls[clave]=xml;
       });
       xmls["xl/workbook.xml"]=forzarRecalculo(xmls["xl/workbook.xml"]);

       var enc=new TextEncoder(),salida={};
       orden.forEach(function(nombre){
         if(xmls[nombre]!==undefined)salida[nombre]={bytes:enc.encode(xmls[nombre])};
         else{
           var e=entradas[nombre];
           salida[nombre]={crudo:e.raw,metodo:e.metodo,crc:e.crc,usize:e.usize};
         }
       });
       return escribirZip(orden,salida).then(function(bytes){
         return {bytes:bytes,escritas:escritas,noEncontradas:noEncontradas,
                 hojasQueFaltan:faltan,partes:orden.length};
       });
     });
  });
}

window.DINXW={aplicar:aplicar,celdasDeRango:celdasDeRango,colNum:colNum,colStr:colStr};
})();


/* ===== p4_views.js ===== */
/* ==========================================================================
   Vistas
   ========================================================================== */
(function(){
"use strict";
var D=window.DIN,V=D.V;
var e=D.esc,eur=D.eur,eur0=D.eur0,eurS=D.eurS,eur0S=D.eur0S,cap=D.cap,n0=D.n0;

/* ------------------------------------------------------------ utilidades */
function cls(){return Array.prototype.filter.call(arguments,Boolean).join(" ");}
function attr(o){
  var s="";for(var k in o){if(o[k]===null||o[k]===undefined||o[k]===false)continue;
    s+=" "+k+'="'+e(o[k])+'"';}
  return s;
}
function num2(v){return v===null||v===undefined||v===""?"":String(Math.round(n0(v)*100)/100);}
function tagCls(c){
  if(c==="PREVISTO"||c==="ANUAL")return "plan";
  if(c==="HUCHA"||c==="INGRESO HUCHA")return "hucha";
  if(c==="AHORRO")return "ahorro";
  if(c==="GASTO HUCHA")return "gasto";
  return "";
}
function tag(c){return '<span class="tag '+tagCls(c)+'">'+e(c||"")+'</span>';}
function dev(v){
  if(v===null||v===undefined)return '<span class="zero">—</span>';
  if(Math.abs(v)<0.005)return '<span class="dev zero">0,00</span>';
  return '<span class="'+cls("dev",v>0?"up":"down")+'">'+eurS(v)+'</span>';
}
function kpi(label,valor,estado,sub,regla){
  return '<div class="kpi"><div class="k-l">'+e(label)+'</div>'+
    '<div class="'+cls("k-n",estado)+'">'+valor+'</div>'+
    (sub?'<div class="k-s">'+sub+'</div>':'')+(regla||'')+'</div>';
}
/* La regla proporcional: el motivo que se repite en toda la app */
function regla(partes){
  var tot=partes.reduce(function(a,p){return a+Math.max(0,p.v);},0);
  if(tot<=0)return '<div class="rule"></div>';
  return '<div class="rule">'+partes.map(function(p){
    var w=Math.max(0,p.v)/tot*100;
    return w<=0?"":'<i class="'+p.c+'" style="width:'+w.toFixed(2)+'%"'+
      (p.t?' data-tip="'+e(p.t)+'"':'')+'></i>';
  }).join("")+'</div>';
}
/* Lista de lo que compone una categoria: apuntes ordenados de mayor a menor, con total */
function listaDesglose(items,vacio,max){
  if(!items.length)return '<div class="desglose"><p class="hint">'+e(vacio)+'</p></div>';
  items=items.slice().sort(function(a,b){return b.importe-a.importe;});
  var tot=items.reduce(function(s,x){return s+x.importe;},0);
  var vis=(max&&items.length>max)?items.slice(0,max):items,resto=items.slice(vis.length);
  return '<div class="desglose">'+vis.map(function(x){
      return '<div class="dg-row"><span class="dg-c">'+e(x.concepto)+(x.sub?'<small>'+e(x.sub)+'</small>':'')+'</span>'+
        '<span class="dg-v">'+eur(x.importe)+'</span></div>';}).join("")+
    (resto.length?'<div class="dg-row"><span class="dg-c hint">… y '+resto.length+' apuntes más</span>'+
      '<span class="dg-v">'+eur(resto.reduce(function(s,x){return s+x.importe;},0))+'</span></div>':'')+
    '<div class="dg-row dg-tot"><span>Total</span><span class="dg-v">'+eur(tot)+'</span></div></div>';
}
/* Medidor de presupuesto. Se toca para ver el desglose de lo que lo compone. */
function medidor(m,C){
  var pres=n0(m.pres),gas=n0(m.gasto);
  var r=pres>0?gas/pres:(gas>0?1.5:0);
  var over=!m.bueno&&!m.hucha&&r>1.0001;
  var w=Math.min(100,r*100);
  var fill=cls("meter-fill",over?"over":"",m.hucha?"hucha":"");
  var restante=pres-gas;
  var abierto=V.desglose===m.k,des="";
  if(abierto)des=listaDesglose(D.desgloseCat(m.cats||[],C.anio,C.mes),
    "Nada anotado en "+m.nom.toLowerCase()+" este mes.");
  return '<div class="tap" data-act="desglose" data-k="'+e(m.k)+'" title="Toca para ver el desglose">'+
    '<div class="meter-top"><b>'+e(m.nom)+'</b><span class="chev'+(abierto?" open":"")+'">›</span>'+
    '<span class="m-v"><em>'+eur(gas)+'</em> / '+eur(pres)+'</span></div>'+
    '<div class="meter-track" data-tip="'+e(m.nom+": "+eur(gas)+" de "+eur(pres))+'">'+
      '<div class="'+fill+'" style="width:'+w.toFixed(2)+'%"></div>'+
      (r>1.0001?'<div class="meter-cap" style="left:calc('+(100/r).toFixed(2)+'% - 1px)"></div>':'')+
    '</div>'+
    '<div class="meter-foot"><span>'+(pres>0?D.pct(r):"—")+'</span>'+
    '<span class="'+(restante<-0.005&&!m.bueno?"over":"")+'">'+
      (restante>=0?"quedan "+eur(restante):(m.bueno?"de sobra "+eur(-restante):"te pasas "+eur(-restante)))+
    '</span></div>'+des+'</div>';
}

/* Techo redondo para los ejes: 1 / 1,5 / 2 / 2,5 / 3 / 4 / 5 / 7,5 x10^n */
function techoBonito(v){
  if(!(v>0))return 1;
  var mag=Math.pow(10,Math.floor(Math.log10(v))),f=v/mag;
  var m=f<=1?1:f<=1.5?1.5:f<=2?2:f<=2.5?2.5:f<=3?3:f<=4?4:f<=5?5:f<=7.5?7.5:10;
  return m*mag;
}
function milesCortos(v){
  v=Math.round(v);
  return v>=1000?(Math.round(v/100)/10).toString().replace(".",",")+"k":String(v);
}

/* ------------------------------------------------------------ graficos */
/* Balance real por mes: polaridad (dos polos + cero neutro), no categorias */
function graficoBalance(R){
  var f=R.filas;
  if(!f.length)return '<p class="hint">Todavía no hay meses en '+R.anio+".</p>";
  var pos=Math.max.apply(null,f.map(function(x){return Math.max(0,x.balanceReal);}).concat([0]));
  var neg=Math.max.apply(null,f.map(function(x){return Math.max(0,-x.balanceReal);}).concat([0]));
  var tot=pos+neg||1,hUp=pos/tot*100,hDn=neg/tot*100;
  var cols=f.map(function(x){
    var v=x.balanceReal;
    var tip=cap(x.mes)+" "+x.anio+": balance "+eurS(v)+" · gastos "+eur0(x.gastoReal)+" · ahorro "+eur0(x.ahoReal);
    return '<div class="col" data-tip="'+e(tip)+'">'+
      '<div class="up" style="height:'+hUp.toFixed(2)+'%">'+
        (v>0&&pos>0?'<i style="height:'+(v/pos*100).toFixed(2)+'%"></i>':'')+
      '</div>'+
      '<div class="dn" style="height:'+hDn.toFixed(2)+'%">'+
        (v<0&&neg>0?'<i class="neg" style="height:'+(-v/neg*100).toFixed(2)+'%"></i>':'')+
      '</div></div>';
  }).join("");
  return '<div class="chart"><div class="bars" style="position:relative">'+
    (pos>0&&neg>0?'<div class="zeroline" style="position:absolute;left:0;right:0;top:'+hUp.toFixed(2)+'%"></div>':'')+
    cols+'</div>'+
    '<div class="bars-axis">'+f.map(function(x){return '<span>'+cap(x.mes).slice(0,3)+'</span>';}).join("")+'</div>'+
    '<div class="legend"><span><i class="a"></i>te sobra dinero</span><span><i class="x"></i>te falta</span></div></div>';
}
/* Ahorro acumulado frente al objetivo: una sola serie, sin leyenda */
function graficoAhorro(R){
  var f=R.filas;if(!f.length)return "";
  var W=680,H=170,ml=46,mr=14,mt=12,mb=22;
  var maxV=techoBonito(Math.max(R.objetivo,Math.max.apply(null,f.map(function(x){return x.acum;})),1));
  var x=function(i){return ml+(f.length<2?0:i*(W-ml-mr)/(f.length-1));};
  var y=function(v){return mt+(1-v/maxV)*(H-mt-mb);};
  var pts=f.map(function(p,i){return x(i)+","+y(p.acum);});
  var area="M"+x(0)+","+y(0)+" L"+pts.join(" L")+" L"+x(f.length-1)+","+y(0)+" Z";
  var line="M"+pts.join(" L");
  var gy=[0,maxV/2,maxV].map(function(v){
    return '<line class="grid" x1="'+ml+'" y1="'+y(v)+'" x2="'+(W-mr)+'" y2="'+y(v)+'"></line>'+
      '<text x="'+(ml-7)+'" y="'+(y(v)+3.2)+'" text-anchor="end">'+milesCortos(v)+'</text>';
  }).join("");
  var goal=R.objetivo>0?'<line class="goal" x1="'+ml+'" y1="'+y(R.objetivo)+'" x2="'+(W-mr)+'" y2="'+y(R.objetivo)+'"></line>'+
      '<text x="'+(W-mr)+'" y="'+(y(R.objetivo)-6)+'" text-anchor="end">objetivo '+milesCortos(R.objetivo)+'</text>':"";
  var hover=f.map(function(p,i){
    var w=(W-ml-mr)/Math.max(1,f.length-1);
    return '<rect x="'+(x(i)-w/2)+'" y="'+mt+'" width="'+w+'" height="'+(H-mt-mb)+'" fill="transparent"'+
      ' data-tip="'+e(cap(p.mes)+": acumulado "+eur0(p.acum)+" · del mes "+eur0(p.ahoReal))+'"></rect>';
  }).join("");
  var dots=f.map(function(p,i){
    return i===f.length-1?'<circle class="end" cx="'+x(i)+'" cy="'+y(p.acum)+'" r="4.5"></circle>':"";
  }).join("");
  return '<div class="chart"><svg class="spark" viewBox="0 0 '+W+' '+H+'" role="img"'+
    ' aria-label="Ahorro acumulado del año frente al objetivo">'+gy+goal+
    '<path class="area" d="'+area+'"></path><path class="line" d="'+line+'"></path>'+dots+hover+
    '</svg><div class="bars-axis" style="padding:0 14px 0 46px">'+
    f.map(function(p){return '<span>'+cap(p.mes).slice(0,3)+'</span>';}).join("")+'</div></div>';
}
/* Gasto por categoria: magnitud -> un solo tono, con marca del previsto.
   Cada categoria se toca para ver los apuntes que la componen. */
function graficoCats(R){
  var cs=R.cats.filter(function(c){return c.real>0.005||c.plan>0.005;});
  if(!cs.length)return '<p class="hint">Sin gasto registrado todavía.</p>';
  cs=cs.slice().sort(function(a,b){return b.real-a.real;});
  var max=Math.max.apply(null,cs.map(function(c){return Math.max(c.real,c.plan);}).concat([1]));
  function des(c){
    if(V.desgloseAnio!==c.cat)return "";
    var items=D.desgloseCat([c.cat],R.anio,null,R.hasta);
    return listaDesglose(items,"Nada anotado en "+cap(c.cat).toLowerCase()+(R.hasta?" hasta "+cap(R.hasta):" en "+R.anio)+".",25);
  }
  if(V.tablaCat){
    return '<div class="tw"><table class="t mid"><thead><tr><th>Categoría</th>'+
      '<th class="n">Real</th><th class="n">Previsto</th><th class="n">Desviación</th></tr></thead><tbody>'+
      cs.map(function(c){
        var d=des(c);
        return '<tr class="tap" data-act="desglose-anio" data-cat="'+e(c.cat)+'"><td>'+tag(c.cat)+'</td><td class="n">'+eur(c.real)+
          '</td><td class="n">'+eur(c.plan)+'</td><td class="n">'+dev(c.real-c.plan)+'</td></tr>'+
          (d?'<tr><td colspan="4" style="padding-top:0">'+d+'</td></tr>':'');}).join("")+
      '</tbody></table></div>';
  }
  return '<div class="hbars">'+cs.map(function(c){
    var tip=c.cat+": real "+eur(c.real)+(c.plan>0?" · previsto "+eur(c.plan):"");
    var d=des(c);
    return '<div class="hbar tap'+(d?" open":"")+'" data-act="desglose-anio" data-cat="'+e(c.cat)+'" data-tip="'+e(tip)+'">'+
      '<div class="hb-l">'+e(cap(c.cat))+'</div>'+
      '<div class="hb-t"><div class="hb-f" style="width:'+(c.real/max*100).toFixed(2)+'%"></div>'+
        (c.plan>0.005?'<div class="hb-g" style="left:calc('+(c.plan/max*100).toFixed(2)+'% - 1px)"></div>':'')+
      '</div><div class="hb-v">'+eur0(c.real)+'</div></div>'+
      (d?'<div class="hbar-des">'+d+'</div>':'');
  }).join("")+'</div>'+
  '<div class="legend"><span><i class="a"></i>gasto real</span><span><i class="g"></i>previsto</span>'+
  '<span class="hint-sm">toca una categoría para ver qué la compone</span></div>';
}

/* ------------------------------------------------------------ vista: MES */
function vistaMes(){
  var ms=D.mesesOrden();
  if(!ms.length)return '<p class="hint">No hay meses cargados.</p>';
  var idx=0;
  ms.forEach(function(m,i){if(m.mes===V.mes&&+m.anio===+V.anio)idx=i;});
  var m=ms[idx];V.mes=m.mes;V.anio=m.anio;
  var C=D.calcMes(m.mes,m.anio);

  var cabecera='<div class="mespick">'+
    '<button class="btn btn-icon" data-act="mes-prev"'+(idx===0?" disabled":"")+' aria-label="Mes anterior">‹</button>'+
    '<button class="btn btn-icon" data-act="mes-next"'+(idx===ms.length-1?" disabled":"")+' aria-label="Mes siguiente">›</button>'+
    '<h1>'+cap(m.mes)+' <i>'+m.anio+'</i></h1>'+
    '<select class="sel" data-act="mes-sel" aria-label="Elegir mes">'+
      ms.map(function(x,i){return '<option value="'+i+'"'+(i===idx?" selected":"")+'>'+
        cap(x.mes)+" "+x.anio+'</option>';}).join("")+
    '</select></div>';

  var kpis='<div class="kpis">'+
    kpi("Balance del mes",eur0S(C.balanceReal),C.balanceReal>=0?"pos":"neg",
      (Math.abs(C.balancePlan)<1?"El plan manda al ahorro todo lo que sobra":
        "El plan dejaba <b>"+eur0S(C.balancePlan)+"</b>"),
      regla([{v:C.gastoReal,c:"a",t:"Gastos "+eur(C.gastoReal)},
             {v:C.ahoReal,c:"o",t:"Ahorro "+eur(C.ahoReal)},
             {v:Math.max(0,C.balanceReal),c:"g",t:"Sin asignar "+eur(Math.max(0,C.balanceReal))}]))+
    kpi("Gastos reales",eur0(C.gastoReal),null,
      "Previstos <b>"+eur0(C.gastoPlan)+"</b> · "+
      (C.gastoReal<=C.gastoPlan?'<span class="pos">'+eur0S(C.gastoReal-C.gastoPlan)+"</span>":
        '<span class="neg">'+eur0S(C.gastoReal-C.gastoPlan)+"</span>"),
      regla([{v:Math.min(C.gastoReal,C.gastoPlan),c:"a"},
             {v:Math.max(0,C.gastoReal-C.gastoPlan),c:"x",t:"Por encima del plan"},
             {v:Math.max(0,C.gastoPlan-C.gastoReal),c:"g",t:"Margen que queda"}]))+
    kpi("Ahorro del mes",eur0(C.ahoReal),C.ahoReal>=C.ahoPlan-0.005?"pos":null,
      "Objetivo <b>"+eur0(C.ahoPlan)+"</b>",
      regla([{v:Math.min(C.ahoReal,C.ahoPlan),c:"a"},
             {v:Math.max(0,C.ahoPlan-C.ahoReal),c:"g",t:"Te falta para el objetivo"}]))+
    '</div>';

  var ingresos='<div class="card card-pad"><div class="sechead"><h2>De dónde sale</h2>'+
    '<span class="right">el neto ya lleva la retención reducida de este mes</span></div>'+
    '<div class="tw"><table class="t"><tbody>'+
    '<tr><td class="c-name">Sueldo neto'+(C.bono>0?" + bono anual":"")+'</td><td class="n">'+eur(C.neto)+'</td></tr>'+
    '<tr class="sub"><td class="c-name">Ticket restaurante (Pluxee)<small>Va aparte: no entra en el balance</small></td>'+
      '<td class="n">'+eur(C.pluxee)+'</td></tr>'+
    '<tr class="sum"><td>Total salidas (gastos + ahorro)</td><td class="n">'+eur(C.salidasReal)+'</td></tr>'+
    '</tbody></table></div></div>';

  var presup='<div class="card card-pad"><div class="sechead"><h2>Cómo van los presupuestos</h2></div>'+
    '<div class="meters">'+C.presup.map(function(m){return medidor(m,C);}).join("")+'</div></div>';

  /* --- anotar + apuntes --- */
  var opts=D.CATS_APUNTE.map(function(c){return '<option value="'+e(c)+'">'+e(cap(c))+'</option>';}).join("");
  var form='<form class="quickadd" data-act="add-apunte">'+
    '<div class="field"><label>Concepto</label><input class="inp" name="concepto" required placeholder="Café 360" autocomplete="off"></div>'+
    '<div class="field"><label>Categoría</label><select class="inp" name="categoria">'+opts+'</select></div>'+
    '<div class="field"><label>Importe</label><input class="inp n" name="importe" type="number" step="0.01" required placeholder="0,00" inputmode="decimal"></div>'+
    '<div class="field"><label>Fecha</label><input class="inp" name="fecha" type="date"></div>'+
    '<div class="field"><label>&nbsp;</label><button class="btn btn-primary" type="submit">Anotar</button></div>'+
    '</form>';
  var listaAp=C.apuntes.length?
    '<div class="tw"><table class="t mid"><thead><tr><th>Concepto</th><th>Categoría</th>'+
    '<th class="n">Importe</th><th>Fecha</th><th></th></tr></thead><tbody>'+
    C.apuntes.map(function(a,i){
      return '<tr><td><input class="inp" data-act="ap-edit" data-i="'+i+'" data-f="concepto" value="'+e(a.concepto)+'"></td>'+
        '<td><select class="inp" data-act="ap-edit" data-i="'+i+'" data-f="categoria">'+
          D.CATS_APUNTE.map(function(c){return '<option value="'+e(c)+'"'+(c===a.categoria?" selected":"")+'>'+e(cap(c))+'</option>';}).join("")+
          (D.CATS_APUNTE.indexOf(a.categoria)<0?'<option value="'+e(a.categoria)+'" selected>'+e(cap(a.categoria))+'</option>':'')+
        '</select></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="ap-edit" data-i="'+i+'" data-f="importe" type="number" step="0.01" value="'+num2(a.importe)+'"></td>'+
        '<td><input class="inp" style="width:132px" data-act="ap-edit" data-i="'+i+'" data-f="fecha" type="date" value="'+e(a.fecha||"")+'"></td>'+
        '<td><button class="btn btn-sm btn-danger" data-act="del-apunte" data-i="'+i+'" aria-label="Borrar apunte">Borrar</button></td></tr>';
    }).join("")+'</tbody></table></div>':
    '<p class="hint">Aún no has anotado nada este mes. Los gastos que anotes aquí descuentan de los presupuestos de arriba.</p>';
  var apuntes='<div class="card card-pad"><div class="sechead"><h2>Lo que has gastado</h2>'+
    '<span class="right">'+C.apuntes.length+' apunte'+(C.apuntes.length===1?"":"s")+'</span></div>'+
    form+'<hr class="divider">'+listaAp+'</div>';

  /* --- plan del mes --- */
  var filas=C.filas.map(function(f){
    var badge='<button class="auto-pill" data-act="auto" data-id="'+e(f.id)+'" title="'+
      (f.auto?"Sigue al plan. Púlsalo para escribir un importe distinto."
             :"Importe puesto a mano. Púlsalo para que vuelva a seguir el plan.")+
      '">'+(f.auto?"auto":"fijo")+'</button>';
    return '<tr><td class="c-name">'+e(f.concepto)+
        (f.nota?'<small>'+e(f.nota)+'</small>':'')+'</td>'+
      '<td>'+tag(f.cat)+'</td>'+
      '<td class="n">'+eur(f.plan)+'</td>'+
      '<td class="n"><div class="realcell">'+
        '<input class="inp n inp-w'+(f.auto?" zero":"")+'" data-act="real" data-id="'+e(f.id)+'"'+
        ' type="number" step="0.01" placeholder="—" value="'+num2(f.real)+'">'+badge+'</div></td>'+
      '<td class="n">'+dev(f.dev)+'</td></tr>';
  }).join("");
  var H=C.hucha;
  if(H.adelanto>0)filas+='<tr><td class="c-name">Adelanto a la hucha'+
    '<small>La hucha se quedó a cero: esto sale de tu cuenta y te vuelve de las próximas asignaciones</small></td>'+
    '<td>'+tag("HUCHA")+'</td><td class="n">'+eur(0)+'</td><td class="n">'+eur(H.adelanto)+'</td>'+
    '<td class="n">'+dev(H.adelanto)+'</td></tr>';
  if(H.recuperado>0)filas+='<tr><td class="c-name">Recuperas lo adelantado a la hucha'+
    '<small>De la asignación de '+eur0(H.ent)+', '+eur0(H.recuperado)+' vuelven a ti: a la hucha entran '+eur0(H.entFis)+'</small></td>'+
    '<td>'+tag("HUCHA")+'</td><td class="n">'+eur(0)+'</td><td class="n">−'+eur(H.recuperado)+'</td>'+
    '<td class="n">'+dev(-H.recuperado)+'</td></tr>';
  var plan='<div class="card card-pad"><div class="sechead"><h2>El plan del mes</h2>'+
    '<span class="right">«auto» sigue al plan · «fijo» es un importe tuyo · pulsa la etiqueta para cambiar</span></div>'+
    '<div class="tw"><table class="t wide"><thead><tr><th>Concepto</th><th>Categoría</th>'+
    '<th class="n">Plan</th><th class="n">Real</th><th class="n">Desviación</th></tr></thead><tbody>'+filas+
    '<tr class="sum"><td colspan="2">Total gastos (sin ahorro)</td><td class="n">'+eur(C.gastoPlan)+
      '</td><td class="n">'+eur(C.gastoReal)+'</td><td class="n">'+dev(C.gastoReal-C.gastoPlan)+'</td></tr>'+
    '<tr><td colspan="2">Ahorro</td><td class="n">'+eur(C.ahoPlan)+'</td><td class="n">'+eur(C.ahoReal)+
      '</td><td class="n">'+dev(C.ahoReal-C.ahoPlan)+'</td></tr>'+
    '<tr class="sum"><td colspan="2">Balance</td><td class="n">'+eurS(C.balancePlan)+'</td><td class="n '+
      (C.balanceReal>=0?"pos":"neg")+'">'+eurS(C.balanceReal)+'</td><td class="n"></td></tr>'+
    '</tbody></table></div></div>';

  var hucha='<div class="card card-pad"><div class="sechead"><h2>La hucha</h2>'+
    '<span class="right">viajes e imprevistos: no es patrimonio</span></div>'+
    '<div class="tw"><table class="t"><tbody>'+
    '<tr class="sub"><td class="c-name">Saldo al empezar '+cap(m.mes)+'</td><td class="n">'+eur(H.saldoIni)+'</td></tr>'+
    '<tr><td class="c-name">Asignación del mes'+
      (H.recuperado>0?'<small>'+eur0(H.recuperado)+' vuelven a ti por lo que adelantaste: a la hucha entran '+eur0(H.entFis)+'</small>':'')+
      '</td><td class="n">'+eur(H.entFis)+'</td></tr>'+
    '<tr><td class="c-name">Gastos pagados con la hucha'+
      (H.adelanto>0?'<small>La hucha solo cubre '+eur0(H.salFis)+'; los otros '+eur0(H.adelanto)+' los adelantas tú</small>':'')+
      '</td><td class="n">'+(H.salFis>0?'<span class="neg">−'+eur(H.salFis)+'</span>':eur(0))+'</td></tr>'+
    '<tr class="sum"><td>Saldo tras '+cap(m.mes)+'</td><td class="n">'+eur(H.saldo)+'</td></tr>'+
    '</tbody></table></div>'+
    (H.adelanto>0?'<div class="note note-warn" style="margin-top:12px"><span class="ic">!</span>'+
      '<div>La hucha no puede bajar de cero, así que este mes adelantas <b>'+eur(H.adelanto)+'</b> de tu cuenta: '+
      'cuentan como gasto del mes. Te vuelven de las próximas asignaciones de '+eur0(n0(D.S.config.fondoHucha))+
      (H.deuda>H.adelanto+0.005?'; en total la hucha te debe <b>'+eur(H.deuda)+'</b>':'')+'.</div></div>':
     H.deuda>0.005?'<div class="note note-warn" style="margin-top:12px"><span class="ic">!</span>'+
      '<div>La hucha te debe <b>'+eur(H.deuda)+'</b> de lo que adelantaste en meses anteriores. '+
      'Se te devuelve de las próximas asignaciones antes de que la hucha vuelva a crecer.</div></div>':
     H.recuperado>0?'<div class="note note-info" style="margin-top:12px"><span class="ic">i</span>'+
      '<div>Este mes recuperas los <b>'+eur(H.recuperado)+'</b> que adelantaste a la hucha: de la asignación de '+
      eur0(H.ent)+' solo entran '+eur0(H.entFis)+', y el resto se queda en tu cuenta.</div></div>':'')+
    '</div>';

  return cabecera+kpis+'<div class="stack-lg" style="margin-top:20px">'+
    presup+apuntes+plan+ingresos+hucha+'</div>';
}

/* ------------------------------------------------------------ vista: ANO */
/* Dos modos: "Hasta hoy" (lo real hasta el mes en curso, con el plan del resto del ano en
   segundo plano) y "Ano completo" (los 12 meses, donde lo que no ha pasado sigue al plan). */
function vistaAnio(){
  var ys=D.anios();
  if(!ys.length)return '<p class="hint">No hay datos.</p>';
  if(ys.indexOf(+V.anioVista)<0)V.anioVista=ys[0];
  var hoy=new Date(),mesHoy=D.MESES[hoy.getMonth()],anioHoy=hoy.getFullYear();
  var modo=V.anioModo||(+V.anioVista<=anioHoy?"ytd":"anual");
  var hasta=null,avisoYtd="";
  if(modo==="ytd")hasta=+V.anioVista<anioHoy?"DICIEMBRE":(+V.anioVista>anioHoy?null:mesHoy);
  var R=D.resumenAnio(V.anioVista,hasta||undefined);
  if(modo==="ytd"&&(!hasta||!R.filas.length)){
    avisoYtd='<p class="hint" style="margin:-4px 0 14px">'+(+V.anioVista>anioHoy?
      'Este año aún no ha empezado, así que se muestra el plan completo.':
      'Todavía no hay meses de '+V.anioVista+' hasta hoy: se muestra el plan completo.')+'</p>';
    R=D.resumenAnio(V.anioVista);modo="anual";
  }
  var ytd=modo==="ytd",hastaTxt=ytd?cap(R.hasta):"",quedan=ytd?R.resto.meses:0;
  var falta=R.objetivo-R.conseguidoReal;

  var cabecera='<div class="mespick"><h1>Año <i>'+R.anio+'</i></h1>'+
    '<select class="sel" data-act="anio-sel" aria-label="Elegir año">'+
    ys.map(function(y){return '<option value="'+y+'"'+(+y===+R.anio?" selected":"")+'>'+y+'</option>';}).join("")+
    '</select></div>'+
    '<div class="modo"><div class="seg" role="tablist">'+
      '<button class="'+(ytd?"on":"")+'" data-act="anio-modo" data-modo="ytd">Hasta hoy</button>'+
      '<button class="'+(ytd?"":"on")+'" data-act="anio-modo" data-modo="anual">Año completo</button></div>'+
    '<span class="hint">'+(ytd?'lo real hasta '+hastaTxt+(quedan?' · el plan de los '+quedan+' meses que quedan, en segundo plano':''):
      'los '+R.meses+' meses: lo que no ha pasado sigue al plan')+'</span></div>'+avisoYtd;

  var kpis='<div class="kpis-4">'+
    kpi("Ingresos netos",eur0(R.tot.neto),null,R.meses+" mes"+(R.meses===1?"":"es")+(ytd?" hasta "+hastaTxt:" en el plan"))+
    kpi("Gastos reales",eur0(R.tot.gastoReal),null,"Previstos <b>"+eur0(R.tot.gastoPlan)+"</b>"+
      (ytd&&quedan?" · quedan <b>"+eur0(R.resto.gastoPlan)+"</b> de plan":""),
      regla([{v:Math.min(R.tot.gastoReal,R.tot.gastoPlan),c:"a"},
             {v:Math.max(0,R.tot.gastoReal-R.tot.gastoPlan),c:"x"},
             {v:Math.max(0,R.tot.gastoPlan-R.tot.gastoReal),c:"g"}]))+
    kpi("Ahorro acumulado",eur0(R.conseguidoReal),R.conseguidoReal>=R.objetivo?"pos":null,
      (ytd?"Objetivo hasta "+hastaTxt+" <b>"+eur0(R.objetivo)+"</b> · año <b>"+eur0(R.objetivoAnual)+"</b>":
           "Objetivo <b>"+eur0(R.objetivo)+"</b>"),
      regla([{v:Math.min(R.conseguidoReal,R.objetivo),c:"a"},
             {v:Math.max(0,R.objetivo-R.conseguidoReal),c:"g"}]))+
    kpi("Balance acumulado",eur0S(R.tot.balanceReal),R.tot.balanceReal>=0?"pos":"neg",
      "Previsto <b>"+eur0S(R.tot.balancePlan)+"</b>")+
    '</div>';

  var gBal='<div class="card card-pad"><div class="sechead"><h2>Balance de cada mes</h2>'+
    '<span class="right">lo que queda tras gastos y ahorro</span></div>'+graficoBalance(R)+'</div>';
  var gAho='<div class="card card-pad"><div class="sechead"><h2>Ahorro acumulado</h2>'+
    '<span class="right">frente al objetivo'+(ytd?' hasta '+hastaTxt:'')+' de '+eur0(R.objetivo)+'</span></div>'+graficoAhorro(R)+'</div>';

  var tabla='<div class="card card-pad"><div class="sechead"><h2>Mes a mes</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="exp-csv-resumen">Descargar CSV</button></span></div>'+
    '<div class="tw"><table class="t wide"><thead><tr><th>Mes</th><th class="n">Ingresos</th>'+
    '<th class="n">Gastos prev.</th><th class="n">Gastos real</th><th class="n">Ahorro</th>'+
    '<th class="n">Balance</th><th class="n">Ahorro acum.</th></tr></thead><tbody>'+
    R.filas.map(function(f){
      return '<tr><td class="c-name"><a href="#" data-act="ir-mes" data-mes="'+e(f.mes)+'" data-anio="'+f.anio+'"'+
        ' style="color:inherit">'+cap(f.mes)+'</a>'+(f.anotado?'':'<small>sin apuntes</small>')+'</td>'+
        '<td class="n">'+eur0(f.neto)+'</td><td class="n">'+eur0(f.gastoPlan)+'</td>'+
        '<td class="n">'+eur0(f.gastoReal)+'</td><td class="n">'+eur0(f.ahoReal)+'</td>'+
        '<td class="n '+(f.balanceReal>=0?"pos":"neg")+'">'+eur0S(f.balanceReal)+'</td>'+
        '<td class="n">'+eur0(f.acum)+'</td></tr>';
    }).join("")+
    '<tr class="sum"><td>Total '+(ytd?'hasta '+hastaTxt:R.anio)+'</td><td class="n">'+eur0(R.tot.neto)+'</td>'+
    '<td class="n">'+eur0(R.tot.gastoPlan)+'</td><td class="n">'+eur0(R.tot.gastoReal)+'</td>'+
    '<td class="n">'+eur0(R.tot.ahoReal)+'</td><td class="n">'+eur0S(R.tot.balanceReal)+'</td>'+
    '<td class="n">'+eur0(R.conseguidoReal)+'</td></tr>'+
    (ytd&&quedan?'<tr class="sub"><td class="c-name">Resto del año<small>'+quedan+' mes'+(quedan===1?"":"es")+' según el plan</small></td>'+
      '<td class="n">'+eur0(R.resto.neto)+'</td><td class="n">'+eur0(R.resto.gastoPlan)+'</td><td class="n zero">—</td>'+
      '<td class="n">'+eur0(R.resto.ahoPlan)+'</td><td class="n">'+eur0S(R.resto.balancePlan)+'</td>'+
      '<td class="n">'+eur0(R.conseguidoReal+R.resto.ahoPlan)+'</td></tr>':'')+
    '</tbody></table></div></div>';

  var cats='<div class="card card-pad"><div class="sechead"><h2>En qué se va el dinero</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="tabla-cat">'+
    (V.tablaCat?"Ver gráfico":"Ver tabla")+'</button></span></div>'+graficoCats(R)+'</div>';

  var conBolsillo=R.filas.some(function(f){return f.huAdel>0||f.huRec>0;});
  function bolsillo(adel,rec){
    var v=rec-adel;
    if(Math.abs(v)<0.005)return '<span class="zero">—</span>';
    return '<span class="'+(v<0?"neg":"pos")+'">'+eurS(v)+'</span>';
  }
  var hu='<div class="card card-pad"><div class="sechead"><h2>La hucha en '+R.anio+'</h2>'+
    '<span class="right">saldo al empezar el año: '+eur(R.saldoIni)+'</span></div>'+
    '<div class="tw"><table class="t mid"><thead><tr><th>Mes</th><th class="n">Asignas</th>'+
    '<th class="n">Gastas</th>'+(conBolsillo?'<th class="n">Tu bolsillo</th>':'')+'<th class="n">Saldo</th></tr></thead><tbody>'+
    R.filas.map(function(f){
      return '<tr><td>'+cap(f.mes)+'</td><td class="n">'+eur(f.huEnt)+'</td>'+
        '<td class="n">'+(f.huSal>0?'<span class="neg">−'+eur(f.huSal)+'</span>':eur(0))+'</td>'+
        (conBolsillo?'<td class="n">'+bolsillo(f.huAdel,f.huRec)+'</td>':'')+
        '<td class="n">'+eur(f.huSaldo)+'</td></tr>';
    }).join("")+
    '<tr class="sum"><td>Total</td><td class="n">'+eur(R.tot.huEnt)+'</td><td class="n">'+eur(R.tot.huSal)+'</td>'+
    (conBolsillo?'<td class="n">'+bolsillo(R.tot.huAdel,R.tot.huRec)+'</td>':'')+
    '<td class="n">'+eur(R.saldoFin)+'</td></tr></tbody></table></div>'+
    (conBolsillo?'<p class="hint" style="margin-top:10px">La hucha no baja de cero: cuando gastas más de lo que tiene, '+
      'lo adelantas tú (negativo en «Tu bolsillo», cuenta como gasto de ese mes) y te vuelve de las asignaciones siguientes '+
      '(positivo: esa parte no llega a entrar en la hucha).'+
      (R.deudaFin>0.005?' A fin de año la hucha aún te debe <b>'+eur(R.deudaFin)+'</b>.':'')+'</p>':'')+
    '</div>';

  var mensual=R.conseguidoReal-R.bono;
  var prevision=R.conseguidoReal+R.resto.ahoPlan,faltaAnual=R.objetivoAnual-prevision;
  /* en "hasta hoy" la hucha que cuenta para el cierre es la prevista a fin de ano, no la de hoy */
  var huchaFin=ytd?D.resumenAnio(V.anioVista).saldoFin:R.saldoFin;
  var cierre='<div class="card card-pad"><div class="sechead"><h2>¿Llegas al objetivo?</h2></div>'+
    '<div class="tw"><table class="t"><tbody>'+
    '<tr><td class="c-name">Ahorro de las mensualidades'+(ytd?' hasta '+hastaTxt:'')+'</td><td class="n">'+eur(mensual)+'</td></tr>'+
    (R.bono>0?'<tr><td class="c-name">Bono anual<small>va directo a inversión, no toca el mes</small></td>'+
      '<td class="n">'+eur(R.bono)+'</td></tr>':'')+
    '<tr class="sum"><td>Ahorrado '+(ytd?'hasta '+hastaTxt:'en '+R.anio)+'</td><td class="n">'+eur(R.conseguidoReal)+'</td></tr>'+
    '<tr class="sub"><td class="c-name">Objetivo'+(ytd?' hasta '+hastaTxt:' del año')+'<small>'+R.meses+' × '+eur0(n0(D.S.config.ahorroObjetivo))+' al mes</small></td>'+
      '<td class="n">'+eur(R.objetivo)+'</td></tr>'+
    '<tr><td class="c-name">'+(falta>0?"Te falta":"Vas por delante")+'</td><td class="n '+(falta>0?"neg":"pos")+'">'+
      eur(Math.abs(falta))+'</td></tr>'+
    (ytd&&quedan?'<tr class="sub"><td class="c-name">Previsión a fin de año<small>si los '+quedan+' meses que quedan van según el plan ('+
        eur0(R.resto.ahoPlan)+' más)</small></td><td class="n">'+eur(prevision)+'</td></tr>'+
      '<tr class="sub"><td class="c-name">Objetivo del año<small>'+R.mesesAnio+' × '+eur0(n0(D.S.config.ahorroObjetivo))+' al mes</small></td>'+
        '<td class="n">'+eur(R.objetivoAnual)+'</td></tr>':'')+
    '<tr class="sub"><td class="c-name">Saldo de la hucha'+(ytd?' hasta '+hastaTxt:' a fin de año')+'<small>no es patrimonio, pero está ahí</small></td>'+
      '<td class="n">'+eur(R.saldoFin)+'</td></tr>'+
    (ytd&&quedan?
      '<tr class="sum"><td>A fin de año, contando también la hucha<small>hucha prevista a fin de año: '+eur(huchaFin)+'</small></td><td class="n '+
        ((faltaAnual-huchaFin)<=0?"pos":"neg")+'">'+
        ((faltaAnual-huchaFin)<=0?"objetivo cubierto":"faltarían "+eur(faltaAnual-huchaFin))+'</td></tr>':
      '<tr class="sum"><td>Contando también la hucha</td><td class="n '+
        ((falta-R.saldoFin)<=0?"pos":"neg")+'">'+
        ((falta-R.saldoFin)<=0?"objetivo cubierto":"faltarían "+eur(falta-R.saldoFin))+'</td></tr>')+
    '</tbody></table></div></div>';

  return cabecera+kpis+'<div class="stack-lg" style="margin-top:20px">'+
    gBal+gAho+tabla+cats+hu+cierre+'</div>';
}

window.DINV1={vistaMes:vistaMes,vistaAnio:vistaAnio,
  kpi:kpi,regla:regla,medidor:medidor,tag:tag,dev:dev,num2:num2,attr:attr,cls:cls};
})();


/* ===== p5x_vistas.js ===== */
/* ==========================================================================
   Vistas de Reparto, Anuales y Ajustes
   ========================================================================== */
(function(){
"use strict";
var D=window.DIN,U=window.DINV1,V=D.V;
var e=D.esc,eur=D.eur,eur0=D.eur0,eurS=D.eurS,eur0S=D.eur0S,cap=D.cap,n0=D.n0;
var kpi=U.kpi,regla=U.regla,tag=U.tag,num2=U.num2;

/* ------------------------------------------------------------ vista: REPARTO */
function vistaReparto(){
  var T=D.repTotales(),c=D.S.config;
  var wD=T.total>0?T.aD/T.total*100:50;
  var cabecera='<div class="mespick"><h1>Gastos <i>comunes</i></h1></div>';

  var kpis='<div class="kpis">'+
    kpi("Total común al mes",eur0(T.total),null,"entre los dos")+
    kpi("Tu parte",eur0(T.aD),null,"el "+D.pct(T.pD)+" del total",
      regla([{v:T.aD,c:"a",t:"David "+eur(T.aD)},{v:T.aC,c:"o",t:"Celine "+eur(T.aC)}]))+
    kpi("A transferir al bote",eur0(T.trD),null,"ya pagas <b>"+eur0(T.directo)+"</b> por tu cuenta")+
    '</div>';

  var split='<div class="card card-pad"><div class="sechead"><h2>Cómo se reparte</h2>'+
    '<span class="right">proporcional a los sueldos netos</span></div>'+
    '<div class="split">'+
      '<div class="d" style="width:'+wD.toFixed(2)+'%" data-tip="'+e("David: "+eur(T.aD)+" · "+D.pct(T.pD))+'">'+
        (wD>18?D.pct(T.pD):"")+'</div>'+
      '<div class="c" style="width:'+(100-wD).toFixed(2)+'%" data-tip="'+e("Celine: "+eur(T.aC)+" · "+D.pct(T.pC))+'">'+
        (100-wD>18?D.pct(T.pC):"")+'</div>'+
    '</div>'+
    '<div class="legend"><span><i class="a"></i>David · '+eur(T.aD)+'</span>'+
      '<span><i class="o"></i>Celine · '+eur(T.aC)+'</span></div>'+
    '<hr class="divider">'+
    '<div class="tw"><table class="t"><tbody>'+
    '<tr><td class="c-name">Sueldo neto de David</td><td class="n">'+
      '<input class="inp n inp-w" data-act="cfg" data-k="sueldoDavid" type="number" step="1" value="'+num2(c.sueldoDavid)+'"></td></tr>'+
    '<tr><td class="c-name">Sueldo neto de Celine</td><td class="n">'+
      '<input class="inp n inp-w" data-act="cfg" data-k="sueldoCeline" type="number" step="1" value="'+num2(c.sueldoCeline)+'"></td></tr>'+
    '<tr><td class="c-name">Tu aportación<small>'+D.pct(T.pD)+' de '+eur(T.total)+'</small></td><td class="n">'+eur(T.aD)+'</td></tr>'+
    '<tr><td class="c-name">Menos lo que ya pagas directamente</td><td class="n">−'+eur(T.directo)+'</td></tr>'+
    '<tr class="sum"><td>Te toca transferir</td><td class="n">'+eur(T.trD)+'</td></tr>'+
    '<tr class="sub"><td class="c-name">Si fuese al 50 / 50<small>pagarías '+eurS(T.difD)+' respecto a lo de arriba</small></td>'+
      '<td class="n">'+eur(T.mitad)+'</td></tr>'+
    '</tbody></table></div></div>';

  var lineas='<div class="card card-pad"><div class="sechead"><h2>Qué entra en el bote</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="rep-add">Añadir línea</button></span></div>'+
    '<div class="tw"><table class="t wide"><thead><tr><th>Concepto</th><th>Detalle</th>'+
    '<th class="n">€ / mes</th><th>Notas</th><th></th></tr></thead><tbody>'+
    T.lineas.map(function(l,i){
      return '<tr><td><input class="inp" data-act="rep" data-i="'+i+'" data-f="grupo" value="'+e(l.grupo)+'"></td>'+
        '<td><input class="inp" data-act="rep" data-i="'+i+'" data-f="subconcepto" value="'+e(l.subconcepto||"")+'"></td>'+
        '<td class="n">'+(l.linkAnual?
          '<span data-tip="'+e("Viene del gasto anual «"+l.linkAnual+"»: "+eur(l.imp*12)+" al año")+'">'+eur(l.imp)+
            '<br><span class="hint-sm">anual / 12</span></span>':
          '<input class="inp n inp-w" data-act="rep" data-i="'+i+'" data-f="importe" type="number" step="0.01" value="'+num2(l.importe)+'">')+'</td>'+
        '<td><input class="inp" data-act="rep" data-i="'+i+'" data-f="notas" value="'+e(l.notas||"")+'"></td>'+
        '<td><button class="btn btn-sm btn-danger" data-act="rep-del" data-i="'+i+'">Quitar</button></td></tr>';
    }).join("")+
    '<tr class="sum"><td colspan="2">Total común al mes</td><td class="n">'+eur(T.total)+'</td><td colspan="2"></td></tr>'+
    '</tbody></table></div>'+
    '<p class="hint-sm" style="margin-top:10px">Las líneas que ponen «anual / 12» salen de la pestaña '+
    'Anuales: cambia allí el importe del año y aquí se actualiza la doceava parte.</p></div>';

  var grupos='<div class="card card-pad"><div class="sechead"><h2>Peso de cada concepto</h2></div>'+
    '<div class="hbars">'+T.grupos.map(function(g){
      var max=T.grupos[0].imp||1;
      return '<div class="hbar" data-tip="'+e(g.grupo+": "+eur(g.imp)+" al mes · "+D.pct(g.pct))+'">'+
        '<div class="hb-l">'+e(cap(g.grupo))+'</div>'+
        '<div class="hb-t"><div class="hb-f" style="width:'+(g.imp/max*100).toFixed(2)+'%"></div></div>'+
        '<div class="hb-v">'+eur0(g.imp)+'</div></div>';
    }).join("")+'</div></div>';

  return cabecera+kpis+'<div class="stack-lg" style="margin-top:20px">'+split+lineas+grupos+'</div>';
}

/* ------------------------------------------------------------ vista: ANUALES */
function vistaAnuales(){
  var ga=D.S.gastosAnuales||[];
  var com=ga.filter(function(g){return g.ambito==="COMÚN"||g.ambito==="COMUN";});
  var per=ga.filter(function(g){return g.ambito==="PERSONAL";});
  var sc=com.reduce(function(a,b){return a+n0(b.importe);},0);
  var sp=per.reduce(function(a,b){return a+n0(b.importe);},0);
  var mesOpts=function(sel){return D.MESES.map(function(m){
    return '<option value="'+m+'"'+(m===sel?" selected":"")+'>'+cap(m)+'</option>';}).join("");};

  var kpis='<div class="kpis">'+
    kpi("Comunes al año",eur0(sc),null,"ya prorrateados en el bote: <b>"+eur0(sc/12)+"</b> al mes")+
    kpi("Personales al año",eur0(sp),null,"caen enteros en su mes")+
    kpi("Todo junto al mes",eur0((sc+sp)/12),null,"si lo repartieses en doce",
      regla([{v:sc,c:"a",t:"Comunes "+eur0(sc)},{v:sp,c:"o",t:"Personales "+eur0(sp)}]))+
    '</div>';

  var tabla='<div class="card card-pad"><div class="sechead"><h2>Gastos que llegan una vez al año</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="ga-add">Añadir gasto</button></span></div>'+
    '<div class="tw"><table class="t wide"><thead><tr><th>Concepto</th><th class="n">Importe</th>'+
    '<th>Mes de cobro</th><th>Ámbito</th><th class="n">/ mes</th><th></th></tr></thead><tbody>'+
    ga.map(function(g,i){
      return '<tr><td><input class="inp" data-act="ga" data-i="'+i+'" data-f="concepto" value="'+e(g.concepto)+'"></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="ga" data-i="'+i+'" data-f="importe" type="number" step="0.01" value="'+num2(g.importe)+'"></td>'+
        '<td><select class="inp" data-act="ga" data-i="'+i+'" data-f="mes">'+mesOpts(g.mes)+'</select></td>'+
        '<td><select class="inp" data-act="ga" data-i="'+i+'" data-f="ambito">'+
          '<option value="COMÚN"'+(g.ambito!=="PERSONAL"?" selected":"")+'>Común</option>'+
          '<option value="PERSONAL"'+(g.ambito==="PERSONAL"?" selected":"")+'>Personal</option></select></td>'+
        '<td class="n">'+eur(n0(g.importe)/12)+'</td>'+
        '<td><button class="btn btn-sm btn-danger" data-act="ga-del" data-i="'+i+'">Quitar</button></td></tr>';
    }).join("")+
    '<tr class="sum"><td>Total</td><td class="n">'+eur(sc+sp)+'</td><td colspan="2"></td>'+
    '<td class="n">'+eur((sc+sp)/12)+'</td><td></td></tr>'+
    '</tbody></table></div>'+
    '<div class="note note-info" style="margin-top:14px"><span class="ic">i</span><div>'+
    '<b>Común</b> quiere decir que ya está prorrateado en el bote de gastos comunes, así que no vuelve a aparecer en el mes. '+
    '<b>Personal</b> aparece solo, entero, en la pestaña Mes del mes que le toca.</div></div></div>';

  var calendario='<div class="card card-pad"><div class="sechead"><h2>En qué mes cae cada uno</h2></div>'+
    '<div class="hbars">'+D.MESES.map(function(m){
      var enMes=ga.filter(function(g){return g.mes===m;});
      var tot=enMes.reduce(function(a,b){return a+n0(b.importe);},0);
      var maxM=Math.max.apply(null,D.MESES.map(function(mm){
        return ga.filter(function(g){return g.mes===mm;}).reduce(function(a,b){return a+n0(b.importe);},0);}).concat([1]));
      if(!enMes.length)return '<div class="hbar"><div class="hb-l">'+cap(m)+'</div>'+
        '<div class="hb-t"></div><div class="hb-v zero">—</div></div>';
      return '<div class="hbar" data-tip="'+e(cap(m)+": "+enMes.map(function(g){
          return g.concepto+" "+eur0(g.importe)+(g.ambito==="PERSONAL"?" (personal)":"");}).join(" · "))+'">'+
        '<div class="hb-l">'+cap(m)+'</div>'+
        '<div class="hb-t"><div class="hb-f" style="width:'+(tot/maxM*100).toFixed(2)+'%"></div></div>'+
        '<div class="hb-v">'+eur0(tot)+'</div></div>';
    }).join("")+'</div></div>';

  return '<div class="mespick"><h1>Gastos <i>anuales</i></h1></div>'+kpis+
    '<div class="stack-lg" style="margin-top:20px">'+tabla+calendario+'</div>';
}

/* ------------------------------------------------------------ vista: AJUSTES */
function vistaAjustes(){
  var c=D.S.config,ys=D.anios();
  var f=function(k,label,paso,nota){
    return '<div class="field"><label>'+e(label)+'</label>'+
      '<input class="inp n" data-act="cfg" data-k="'+k+'" type="number" step="'+(paso||"1")+'" value="'+num2(c[k])+'">'+
      (nota?'<span class="hint-sm">'+e(nota)+'</span>':'')+'</div>';
  };
  var params='<div class="card card-pad"><div class="sechead"><h2>Los números de partida</h2>'+
    '<span class="right">todo lo demás se calcula solo</span></div><div class="grid2">'+
    f("sueldoDavid","Sueldo neto David","1","Se usa cuando el mes no tiene un neto excepcional")+
    f("sueldoCeline","Sueldo neto Celine","1","Solo para calcular el reparto")+
    f("pluxee","Ticket restaurante","1","Informativo: no entra en el balance")+
    f("presupCompras","Presupuesto compras / mes","5")+
    f("presupVidaDiaria","Presupuesto vida diaria / mes","5")+
    f("ahorroObjetivo","Ahorro objetivo / mes","10","El plan reparte el objetivo anual entre los meses que queden")+
    f("fondoHucha","Fondo hucha / mes","10","Viajes e imprevistos")+
    f("haciendaDefault","Provisión Hacienda por defecto","5","Para los meses sin excepción")+
    f("bonoAnual","Bono anual neto","50")+
    '<div class="field"><label>Mes de cobro del bono</label>'+
      '<select class="inp" data-act="cfg-mesbono">'+
      '<option value="">— sin bono —</option>'+
      D.MESES.map(function(m){return '<option value="'+m+'"'+(m===c.mesBono?" selected":"")+'>'+cap(m)+'</option>';}).join("")+
      '</select></div>'+
    '<div class="field"><label>Año del bono</label>'+
      '<select class="inp" data-act="cfg-aniobono">'+
      ys.map(function(y){return '<option value="'+y+'"'+(+y===+c.anioBono?" selected":"")+'>'+y+'</option>';}).join("")+
      '</select></div>'+
    '</div></div>';

  var curso='<div class="card card-pad"><div class="sechead"><h2>Curso de alemán</h2>'+
    '<span class="right">se anota como Cultura en los meses marcados</span></div>'+
    ys.map(function(y){
      var lst=(D.S.mesesCurso||{})[String(y)]||[];
      return '<div style="margin-bottom:16px"><div class="field" style="max-width:200px;margin-bottom:10px">'+
        '<label>Cuota mensual '+y+'</label>'+
        '<input class="inp n" data-act="curso-imp" data-anio="'+y+'" type="number" step="1" value="'+
          num2((c.cursoAleman||{})[String(y)])+'"></div>'+
        '<div class="nav" style="flex-wrap:wrap;gap:5px">'+D.MESES.map(function(m){
          var on=lst.indexOf(m)>=0;
          return '<button class="btn btn-sm" data-act="curso-mes" data-anio="'+y+'" data-mes="'+m+'"'+
            (on?' aria-current="page" style="border-color:var(--accent);color:var(--accent);font-weight:600"':'')+
            '>'+cap(m).slice(0,3)+'</button>';
        }).join("")+'</div></div>';
    }).join("")+
    (c.cobeeAleman>0?'<div class="note note-info"><span class="ic">i</span><div>El curso se paga vía Cobee: '+
      'el coste neto estimado por convocatoria era <b>'+eur(c.cobeeAleman)+'</b> una vez descontado el ahorro de IRPF. '+
      'Aquí se refleja la cuota visible del mes, no ese efecto en nómina.</div></div>':'')+
    '</div>';

  var exc='<div class="card card-pad"><div class="sechead"><h2>Meses con neto distinto</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="exc-add">Añadir mes</button></span></div>'+
    '<div class="tw"><table class="t wide"><thead><tr><th>Mes</th><th class="n">Año</th>'+
    '<th class="n">Neto</th><th class="n">Provisión Hacienda</th><th></th></tr></thead><tbody>'+
    (D.S.excepciones||[]).map(function(x,i){
      return '<tr><td><select class="inp" data-act="exc" data-i="'+i+'" data-f="mes">'+
        D.MESES.map(function(m){return '<option value="'+m+'"'+(m===x.mes?" selected":"")+'>'+cap(m)+'</option>';}).join("")+
        '</select></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="exc" data-i="'+i+'" data-f="anio" type="number" step="1" value="'+num2(x.anio)+'"></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="exc" data-i="'+i+'" data-f="neto" type="number" step="1" value="'+num2(x.neto)+'"></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="exc" data-i="'+i+'" data-f="hacienda" type="number" step="1" value="'+num2(x.hacienda)+'"></td>'+
        '<td><button class="btn btn-sm btn-danger" data-act="exc-del" data-i="'+i+'">Quitar</button></td></tr>';
    }).join("")+'</tbody></table></div>'+
    '<p class="hint-sm" style="margin-top:10px">Mientras la retención de IRPF esté reducida, el neto sube y hay que '+
    'apartar una provisión para Hacienda. Cuando termine, borra la fila y volverá a usarse el sueldo de arriba.</p></div>';

  var rec='<div class="card card-pad"><div class="sechead"><h2>Gastos fijos de cada mes</h2>'+
    '<span class="right"><button class="btn btn-sm" data-act="rec-add">Añadir</button></span></div>'+
    '<div class="tw"><table class="t mid"><thead><tr><th>Concepto</th><th class="n">€ / mes</th><th></th></tr></thead><tbody>'+
    (D.S.recurrentes||[]).map(function(r,i){
      return '<tr><td><input class="inp" data-act="rec" data-i="'+i+'" data-f="concepto" value="'+e(r.concepto)+'"></td>'+
        '<td class="n"><input class="inp n inp-w" data-act="rec" data-i="'+i+'" data-f="importe" type="number" step="0.01" value="'+num2(r.importe)+'"></td>'+
        '<td><button class="btn btn-sm btn-danger" data-act="rec-del" data-i="'+i+'">Quitar</button></td></tr>';
    }).join("")+'</tbody></table></div>'+
    '<p class="hint-sm" style="margin-top:10px">Aparecen en todos los meses con este importe. '+
    'Los que cambian de un mes a otro es mejor anotarlos como gasto suelto.</p></div>';

  var ms=D.mesesOrden();
  var meses='<div class="card card-pad"><div class="sechead"><h2>Meses del plan</h2>'+
    '<span class="right">'+ms.length+' meses</span></div>'+
    '<div class="nav" style="flex-wrap:wrap;gap:5px;margin-bottom:14px">'+
    ms.map(function(m){return '<span class="btn btn-sm" style="cursor:default">'+cap(m.mes).slice(0,3)+" "+
      String(m.anio).slice(2)+' <button data-act="mes-del" data-mes="'+e(m.mes)+'" data-anio="'+m.anio+
      '" aria-label="Quitar '+cap(m.mes)+" "+m.anio+'" style="color:var(--oxide);font-weight:600">×</button></span>';}).join("")+
    '</div><form class="quickadd" data-act="mes-add" style="grid-template-columns:1fr 1fr auto">'+
    '<div class="field"><label>Mes</label><select class="inp" name="mes">'+
      D.MESES.map(function(m){return '<option value="'+m+'">'+cap(m)+'</option>';}).join("")+'</select></div>'+
    '<div class="field"><label>Año</label><input class="inp n" name="anio" type="number" step="1" value="'+
      (ms.length?ms[ms.length-1].anio:new Date().getFullYear())+'"></div>'+
    '<div class="field"><label>&nbsp;</label><button class="btn" type="submit">Añadir mes</button></div>'+
    '</form>'+
    '<p class="hint-sm" style="margin-top:10px">El objetivo de ahorro anual se reparte entre los meses que tengas '+
    'aquí de cada año, así que si añades o quitas meses el plan de ahorro se recalcula.</p></div>';

  var excel=window.DINPWA.tarjetaCuenta();

  return '<div class="mespick"><h1>Ajustes</h1></div>'+
    '<div class="stack-lg">'+params+curso+exc+rec+meses+excel+'</div>';
}

window.DINV2={vistaReparto:vistaReparto,vistaAnuales:vistaAnuales,vistaAjustes:vistaAjustes};
})();


/* ===== p6_pwa.js ===== */
/* ==========================================================================
   Capa de plataforma del PWA: cuenta Microsoft y lectura/escritura del Excel
   en OneDrive por medio de Microsoft Graph. Sin librerias externas.
   ========================================================================== */
(function(){
"use strict";
var D=window.DIN,U=window.DINV1,U2=window.DINV2,V=D.V;
var e=D.esc,eur=D.eur,eur0=D.eur0,cap=D.cap,n0=D.n0;
var kpi=U.kpi,regla=U.regla,num2=U.num2;

var VERSION="cb3e9d4";
var LSK="dineritos.estado.v2";     /* el estado de trabajo */
var LSC="dineritos.cuenta.v2";     /* clientId, sesion y fichero elegido */
var LSB="dineritos.base.v2";       /* copia de lo ultimo leido del Excel */
var AUT="https://login.microsoftonline.com/consumers/oauth2/v2.0";
var GRAPH="https://graph.microsoft.com/v1.0";
var SCOPES="openid profile offline_access Files.ReadWrite";

var C={clientId:"",token:"",exp:0,refresh:"",cuenta:"",drive:"",item:"",nombre:"",
       etag:"",leido:null};
var BASE=null;        /* estado tal y como estaba en el Excel al leerlo */
var OCUPADO="";       /* texto de la operacion en curso */
var NAV={modo:"carpeta",ruta:[],carpetas:[],archivos:[],aviso:""};  /* explorador de OneDrive */

/* ---------------------------------------------------------------- almacen */
function cargarLocal(){
  try{
    var c=JSON.parse(localStorage.getItem(LSC)||"null");
    if(c)for(var k in c)if(k in C)C[k]=c[k];
  }catch(x){}
  var s=null;
  try{s=JSON.parse(localStorage.getItem(LSK)||"null");}catch(x){}
  /* El estado nunca puede quedarse en null: las vistas leen S.config directamente */
  D.S=(s&&s.meses)?s:D.estadoVacio();
  try{BASE=JSON.parse(localStorage.getItem(LSB)||"null");}catch(x){BASE=null;}
}
function guardarCuenta(){try{localStorage.setItem(LSC,JSON.stringify(C));}catch(x){}}
function guardarLocalEstado(){try{localStorage.setItem(LSK,JSON.stringify(D.S));}catch(x){}}
function fijarBase(){
  BASE=JSON.parse(JSON.stringify(D.S));
  try{localStorage.setItem(LSB,JSON.stringify(BASE));}catch(x){}
}
function conectado(){return !!(C.refresh||(C.token&&Date.now()<C.exp));}
function listo(){return conectado()&&C.drive&&C.item;}

/* ---------------------------------------------------------------- PKCE */
function b64url(buf){
  var b=new Uint8Array(buf),s="";
  for(var i=0;i<b.length;i++)s+=String.fromCharCode(b[i]);
  return btoa(s).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");
}
function aleatorio(n){var a=new Uint8Array(n);crypto.getRandomValues(a);return b64url(a);}
function sha256(t){return crypto.subtle.digest("SHA-256",new TextEncoder().encode(t));}
/* La URI de redireccion tiene que coincidir EXACTAMENTE con la registrada,
   asi que se normaliza: tanto /dineritos/ como /dineritos/index.html dan lo mismo. */
function volverA(){return location.origin+location.pathname.replace(/index\.html$/i,"");}

function entrar(){
  if(!C.clientId){aviso("Antes necesito el ID de aplicación (está en Ajustes).");return;}
  if(!crypto.subtle){aviso("El navegador necesita HTTPS para iniciar sesión. Abre la app por su dirección https.");return;}
  var ver=aleatorio(48),est=aleatorio(16);
  try{
    /* en localStorage y no en sessionStorage: en el movil la vuelta del login
       puede aterrizar en otra pestana o incluso en otro navegador */
    localStorage.setItem("din.pkce",ver);
    localStorage.setItem("din.estado",est);
  }catch(x){aviso("No puedo guardar la sesión: ¿estás en modo privado?");return;}
  sha256(ver).then(function(h){
    location.href=AUT+"/authorize?client_id="+encodeURIComponent(C.clientId)+
      "&response_type=code&redirect_uri="+encodeURIComponent(volverA())+
      "&scope="+encodeURIComponent(SCOPES)+
      "&code_challenge="+b64url(h)+"&code_challenge_method=S256"+
      "&state="+est+"&prompt=select_account";
  });
}
function salir(){
  C.token="";C.refresh="";C.exp=0;C.cuenta="";
  guardarCuenta();pintar();
  aviso("Sesión cerrada. Tus datos siguen en este dispositivo.");
}
function guardarSesion(j){
  if(!j||!j.access_token)throw new Error((j&&(j.error_description||j.error))||"respuesta de sesión vacía");
  C.token=j.access_token;
  C.exp=Date.now()+(((+j.expires_in)||3600)-90)*1000;
  if(j.refresh_token)C.refresh=j.refresh_token;
  if(j.id_token){
    try{
      var p=JSON.parse(decodeURIComponent(atob(j.id_token.split(".")[1].replace(/-/g,"+").replace(/_/g,"/"))
        .split("").map(function(ch){return "%"+("00"+ch.charCodeAt(0).toString(16)).slice(-2);}).join("")));
      C.cuenta=p.preferred_username||p.email||p.name||"";
    }catch(x){}
  }
  guardarCuenta();
}
function pedirToken(cuerpo){
  return fetch(AUT+"/token",{method:"POST",
    headers:{"Content-Type":"application/x-www-form-urlencoded"},
    body:new URLSearchParams(cuerpo)})
   .then(function(r){return r.json();})
   .then(function(j){
     if(j.error){
       var err=new Error(j.error_description||j.error);err.codigo=j.error;throw err;
     }
     guardarSesion(j);return C.token;
   });
}
function procesarVuelta(){
  var q=new URLSearchParams(location.search);
  if(q.get("error")){
    var msg=q.get("error_description")||q.get("error");
    history.replaceState({},"",volverA());
    return Promise.resolve({error:msg});
  }
  var code=q.get("code");
  if(!code)return Promise.resolve(null);
  var est=q.get("state"),guardado=null,ver=null;
  try{
    guardado=localStorage.getItem("din.estado");
    ver=localStorage.getItem("din.pkce");
    localStorage.removeItem("din.estado");localStorage.removeItem("din.pkce");
  }catch(x){}
  history.replaceState({},"",volverA());
  if(!ver)return Promise.resolve({error:"Se ha perdido el hilo del login. Si abriste la app "+
    "desde el icono de la pantalla de inicio, prueba a entrar la primera vez desde el navegador."});
  if(est!==guardado)return Promise.resolve({error:"La vuelta del login no cuadra; inténtalo otra vez."});
  return pedirToken({client_id:C.clientId,grant_type:"authorization_code",code:code,
    redirect_uri:volverA(),code_verifier:ver,scope:SCOPES})
   .then(function(){return {ok:true};},function(err){return {error:err.message};});
}
function token(){
  if(C.token&&Date.now()<C.exp)return Promise.resolve(C.token);
  if(!C.refresh)return Promise.reject(new Error("No hay sesión: entra con tu cuenta de Microsoft."));
  return pedirToken({client_id:C.clientId,grant_type:"refresh_token",
    refresh_token:C.refresh,scope:SCOPES})
   .catch(function(err){
     C.refresh="";C.token="";guardarCuenta();
     throw new Error("La sesión ha caducado (pasa cada 24 h). Vuelve a entrar y no perderás nada.");
   });
}

/* ---------------------------------------------------------------- Graph */
function g(ruta,opts){
  return token().then(function(t){
    opts=opts||{};
    var h={Authorization:"Bearer "+t};
    for(var k in (opts.headers||{}))h[k]=opts.headers[k];
    return fetch(GRAPH+ruta,{method:opts.method||"GET",headers:h,body:opts.body})
     .then(function(r){
       return r.text().then(function(txt){
         var j=null;try{j=txt?JSON.parse(txt):null;}catch(x){}
         if(!r.ok){
           var m=(j&&j.error&&j.error.message)||(r.status+" "+r.statusText);
           var err=new Error(m);err.status=r.status;throw err;
         }
         return j;
       });
     });
  });
}
function ruta(){return "/drives/"+encodeURIComponent(C.drive)+"/items/"+encodeURIComponent(C.item);}

var CAMPOS="id,name,folder,file,parentReference,lastModifiedDateTime,size";
function esLibro(x){return !x.folder&&/\.xls[xm]$/i.test(x.name||"");}
function reparte(items,aviso){
  NAV.carpetas=items.filter(function(x){return !!x.folder;});
  NAV.archivos=items.filter(esLibro);
  NAV.aviso=aviso||"";
  return NAV;
}
/* El explorador no depende del indice de busqueda, que tarda en ver un archivo
   recien subido. Es la via fiable. */
function verCarpeta(itemId,nombre){
  var url=itemId?("/me/drive/items/"+encodeURIComponent(itemId)+"/children")
                :"/me/drive/root/children";
  return g(url+"?$top=200&$select="+CAMPOS).then(function(j){
    NAV.modo="carpeta";
    if(itemId===null)NAV.ruta=[];
    else if(nombre!==undefined)NAV.ruta=NAV.ruta.concat([{id:itemId,nombre:nombre}]);
    var items=(j&&j.value)||[];
    items.sort(function(a,b){return String(a.name||"").localeCompare(String(b.name||""),"es");});
    return reparte(items,items.length>=200?"Solo muestro los 200 primeros de esta carpeta.":"");
  });
}
function subirNivel(indice){
  /* indice -1 = raiz; si no, hasta esa miga */
  if(indice<0){NAV.ruta=[];return verCarpeta(null);}
  var destino=NAV.ruta[indice];
  NAV.ruta=NAV.ruta.slice(0,indice+1);
  return g("/me/drive/items/"+encodeURIComponent(destino.id)+"/children?$top=200&$select="+CAMPOS)
   .then(function(j){
     NAV.modo="carpeta";
     var items=(j&&j.value)||[];
     items.sort(function(a,b){return String(a.name||"").localeCompare(String(b.name||""),"es");});
     return reparte(items,"");
   });
}
function verRecientes(){
  return g("/me/drive/recent?$top=50").then(function(j){
    NAV.modo="recientes";NAV.ruta=[];
    var items=((j&&j.value)||[]).filter(function(x){return !x.folder;});
    return reparte(items,"Lo que has abierto o subido últimamente. Aquí sí aparece un archivo "+
      "recién subido, aunque la búsqueda todavía no lo vea.");
  });
}
function buscarLibro(q){
  var t=q.replace(/'/g,"''");
  return g("/me/drive/root/search(q='"+encodeURIComponent(t)+"')?$top=200&$select="+CAMPOS)
   .then(function(j){
     NAV.modo="busqueda";NAV.ruta=[];
     var todos=(j&&j.value)||[];
     var libros=todos.filter(esLibro);
     return reparte(todos, libros.length?"":
       "La búsqueda no ha encontrado ningún .xlsx con ese nombre. Si acabas de subir el archivo, "+
       "el índice de OneDrive tarda un rato: úsalo desde «Mi OneDrive» o «Recientes».");
   });
}
function metadatos(){
  return g(ruta()+"?$select=id,name,eTag,lastModifiedDateTime,size");
}

function leerExcel(){
  return metadatos().then(function(meta){
    return token().then(function(t){
      return fetch(GRAPH+ruta()+"/content",{headers:{Authorization:"Bearer "+t}})
       .then(function(r){
         if(!r.ok)throw new Error("No he podido descargar el libro ("+r.status+").");
         return r.blob();
       })
       .then(function(b){return D.abrirXlsx(new File([b],meta.name||"libro.xlsx"));})
       .then(function(st){
         st.origen=meta.name||"OneDrive";
         D.S=st;
         C.nombre=meta.name||C.nombre;
         C.etag=meta.eTag||"";
         C.leido=new Date().toISOString();
         guardarCuenta();guardarLocalEstado();fijarBase();
         V.dirty=false;V.mes=null;V.anio=null;V.anioVista=null;
         return st;
       });
    });
  });
}

/* ------------------------- bajar y subir el libro -------------------------- */
function descargarContenido(){
  return token().then(function(t){
    return fetch(GRAPH+ruta()+"/content",{headers:{Authorization:"Bearer "+t}})
     .then(function(r){
       if(!r.ok)throw new Error("No he podido descargar el libro ("+r.status+").");
       return r.arrayBuffer();
     });
  });
}
function subirContenido(bytes,etag){
  return token().then(function(t){
    var h={Authorization:"Bearer "+t,
      "Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};
    if(etag)h["if-match"]=etag;
    return fetch(GRAPH+ruta()+"/content",{method:"PUT",headers:h,body:bytes})
     .then(function(r){
       if(r.status===412||r.status===409){
         var e2=new Error("conflicto");e2.conflicto=true;throw e2;
       }
       return r.text().then(function(txt){
         var j=null;try{j=txt?JSON.parse(txt):null;}catch(x){}
         if(!r.ok){
           var m=(j&&j.error&&j.error.message)||(r.status+" "+r.statusText);
           throw new Error(m);
         }
         return j;
       });
     });
  });
}


/* Valor que le toca a la columna E de una fila de plan: numero, formula o vacio */
function celdaReal(m,fila){
  var id=D.FILAS_FIJAS[fila];
  if(!id)return "";
  var v=(m.reales||{})[id];
  if(v===undefined)v=(id==="COMPRAS"||id==="VIDADIARIA")?null:"auto";
  if(v==="auto")return "=D"+fila;
  if(v===null||v==="")return "";
  return n0(v);
}
function descriptoresMes(m,parte){
  var hoja=m.hojaOrigen;
  if(!hoja)return [];
  var t=[];
  if(parte.apuntes){
    /* Solo B, C, E, G y H. Nunca D ni F: ahi viven formulas de la hoja. */
    var bc=[],ee=[],gh=[];
    for(var k=0;k<(D.AP_FIN-D.AP_INI+1);k++){
      var a=m.apuntes[k];
      bc.push([a?a.concepto:"", a?a.categoria:""]);
      ee.push([a?n0(a.importe):""]);
      gh.push([a&&a.fecha?a.fecha:"", a&&a.notas?a.notas:""]);
    }
    t.push(["B"+D.AP_INI+":C"+D.AP_FIN,{values:bc}]);
    t.push(["E"+D.AP_INI+":E"+D.AP_FIN,{values:ee}]);
    t.push(["G"+D.AP_INI+":H"+D.AP_FIN,{values:gh}]);
  }
  if(parte.reales){
    var f1=[];
    for(var r=6;r<=16;r++)f1.push([celdaReal(m,r)]);
    t.push(["E6:E16",{formulas:f1}]);
    t.push(["E18:E18",{formulas:[[celdaReal(m,18)]]}]);
    /* gastos anuales personales del mes: filas 39-43, en el orden de la hoja */
    var idx=[];
    (D.S.gastosAnuales||[]).forEach(function(x,i){
      if(x.ambito==="PERSONAL"&&x.mes===m.mes)idx.push(i);
    });
    if(idx.length){
      var fg=[];
      for(var k2=0;k2<(D.GA_FIN-D.GA_INI+1);k2++){
        var gi=idx[k2];
        if(gi===undefined){fg.push([""]);continue;}
        var v=(m.reales||{})["GA:"+gi];
        fg.push([v==="auto"||v===undefined?"":(v===null?"":n0(v))]);
      }
      t.push(["E"+D.GA_INI+":E"+D.GA_FIN,{formulas:fg}]);
    }
  }
  return t.map(function(par){
    return {hoja:hoja,dir:par[0],cuerpo:par[1],
            que:cap(m.mes)+" "+m.anio+(par[0].charAt(0)==="B"||par[0].charAt(0)==="G"?" · apuntes":
                (par[0].indexOf("E"+D.AP_INI)===0?" · importes de los apuntes":" · importes reales"))};
  });
}

/* El plan completo de escritura, inspeccionable antes de enviarlo */
function planEscritura(p){
  var plan=[];
  p.meses.forEach(function(x){
    plan=plan.concat(descriptoresMes(x.m,{apuntes:x.apuntes,reales:x.reales}));
  });
  p.config.forEach(function(c){
    plan.push({hoja:"CONFIG",dir:c.ref+":"+c.ref,cuerpo:{values:[[c.v]]},que:"CONFIG · "+c.k});
  });
  p.excepciones.forEach(function(x){
    plan.push({hoja:"CONFIG",dir:"I"+x._fila+":J"+x._fila,
      cuerpo:{values:[[n0(x.neto),n0(x.hacienda)]]},
      que:"CONFIG · neto y Hacienda de "+cap(x.mes)+" "+x.anio});
  });
  p.reparto.forEach(function(l){
    plan.push({hoja:"REPARTO",dir:"D"+l._fila+":D"+l._fila,
      cuerpo:{values:[[n0(l.importe)]]},
      que:"REPARTO · "+l.grupo+" / "+(l.subconcepto||"")});
  });
  p.anuales.forEach(function(x){
    plan.push({hoja:"GASTOS_ANUALES",dir:"B"+x._fila+":E"+x._fila,
      cuerpo:{values:[[x.concepto,n0(x.importe),x.mes,x.ambito]]},
      que:"GASTOS_ANUALES · "+x.concepto});
  });
  return plan;
}

/* Muestra lo que se va a escribir, sin enviar nada */
function simular(){
  var p=pendientes();
  if(!p){aviso("Antes tengo que leer el Excel una vez.");return;}
  if(!p.total){hoja("Nada que escribir",'<p class="hint">El Excel ya está al día.</p>');return;}
  var plan=planEscritura(p);
  var filas=plan.map(function(d){
    var c=d.cuerpo.values||d.cuerpo.formulas;
    var noVacias=c.filter(function(f){
      return f.some(function(v){return v!==""&&v!==null&&v!==undefined;});
    }).length;
    return '<tr><td class="c-name">'+e(d.que)+'</td><td><span class="tag">'+e(d.hoja)+'</span></td>'+
      '<td class="n">'+e(d.dir)+'</td><td class="n">'+noVacias+' / '+c.length+'</td>'+
      '<td>'+(d.cuerpo.formulas?'<span class="tag plan">fórmulas</span>':'<span class="tag">valores</span>')+'</td></tr>';
  }).join("");
  hoja("Lo que voy a escribir",
    '<p class="hint">'+plan.length+' bloques de celdas. Nunca toco las columnas D ni F de las hojas '+
    'mensuales, que es donde viven las fórmulas del plan y la desviación, ni ninguna otra celda '+
    'del libro.</p>'+
    '<div class="tw" style="margin-top:12px"><table class="t"><thead><tr><th>Qué</th><th>Hoja</th>'+
    '<th class="n">Rango</th><th class="n">Celdas con dato</th><th>Tipo</th></tr></thead><tbody>'+
    filas+'</tbody></table></div>'+
    (p.sinSitio.length?'<div class="note note-warn" style="margin-top:12px"><span class="ic">!</span><div>'+
      'Esto no cabe en el Excel y habría que añadirlo a mano:<br>'+
      p.sinSitio.map(function(s){return "· "+e(s);}).join("<br>")+'</div></div>':''),
    '<button class="btn btn-primary" data-act="guardar-excel">Escribirlo</button>'+
    '<button class="btn" data-act="cerrar">Dejarlo</button>');
}

/* Que ha cambiado respecto a lo que leimos del Excel */
function pendientes(){
  if(!BASE)return null;
  var out={meses:[],config:[],reparto:[],anuales:[],excepciones:[],sinSitio:[]};
  var J=function(x){return JSON.stringify(x);};
  (D.S.meses||[]).forEach(function(m){
    var b=(BASE.meses||[]).filter(function(x){return x.mes===m.mes&&+x.anio===+m.anio;})[0];
    if(!b){out.sinSitio.push(cap(m.mes)+" "+m.anio+" (mes nuevo, no existe su hoja)");return;}
    var ap=J(m.apuntes)!==J(b.apuntes), re=J(m.reales)!==J(b.reales);
    if(ap||re)out.meses.push({m:m,apuntes:ap,reales:re});
    if((m.apuntes||[]).length>(D.AP_FIN-D.AP_INI+1))
      out.sinSitio.push(cap(m.mes)+" "+m.anio+": "+m.apuntes.length+" apuntes y en la hoja caben "+
        (D.AP_FIN-D.AP_INI+1));
  });
  var cb=BASE.config||{},cs=D.S.config||{};
  [["F5","sueldoDavid"],["F6","sueldoCeline"],["F7","pluxee"],["F8","presupCompras"],
   ["F9","presupVidaDiaria"],["F10","ahorroObjetivo"],["F12","fondoHucha"],["F13","bonoAnual"],
   ["F14","mesBono"]].forEach(function(par){
    if(J(cs[par[1]])!==J(cb[par[1]]))out.config.push({ref:par[0],k:par[1],v:cs[par[1]]});
  });
  var cua=cs.cursoAleman||{},cub=cb.cursoAleman||{};
  var anios=Object.keys(cua).map(Number).sort(function(a,b2){return a-b2;});
  ["F16","F18"].forEach(function(ref,i){
    var y=anios[i];if(y===undefined)return;
    if(J(cua[String(y)])!==J(cub[String(y)]))out.config.push({ref:ref,k:"curso "+y,v:cua[String(y)]});
  });
  (D.S.excepciones||[]).forEach(function(x){
    var b=(BASE.excepciones||[]).filter(function(y){return y.mes===x.mes&&+y.anio===+x.anio;})[0];
    if(!x._fila){out.sinSitio.push("Excepción de "+cap(x.mes)+" "+x.anio+" (nueva)");return;}
    if(!b||n0(b.neto)!==n0(x.neto)||n0(b.hacienda)!==n0(x.hacienda))out.excepciones.push(x);
  });
  (D.S.reparto||[]).forEach(function(l){
    if(!l._fila){out.sinSitio.push("Línea de reparto «"+l.grupo+" / "+(l.subconcepto||"")+"» (nueva)");return;}
    if(l.linkAnual)return;
    var b=(BASE.reparto||[]).filter(function(y){return y._fila===l._fila;})[0];
    if(!b||n0(b.importe)!==n0(l.importe))out.reparto.push(l);
  });
  (D.S.gastosAnuales||[]).forEach(function(x){
    if(!x._fila){out.sinSitio.push("Gasto anual «"+x.concepto+"» (nuevo)");return;}
    var b=(BASE.gastosAnuales||[]).filter(function(y){return y._fila===x._fila;})[0];
    if(!b||n0(b.importe)!==n0(x.importe)||b.ambito!==x.ambito||b.concepto!==x.concepto)out.anuales.push(x);
  });
  out.total=out.meses.length+out.config.length+out.reparto.length+out.anuales.length+out.excepciones.length;
  return out;
}

function guardarEnExcel(forzar){
  if(!listo()){aviso("Primero conecta tu cuenta y elige el libro (Ajustes).");return;}
  if(!navigator.onLine){
    aviso("Sin conexión. Lo tienes guardado aquí; dale a Guardar cuando vuelvas a tener red.");return;}
  var p=pendientes();
  if(!p){aviso("Antes tengo que leer el Excel una vez para saber de dónde partimos.");return;}
  if(!p.total){V.dirty=false;guardarLocalEstado();pintar();aviso("El Excel ya está al día.");return;}
  var plan=planEscritura(p),info=null;
  ocupar("Bajando el libro…");
  metadatos()
   .then(function(meta){
     /* se escribe siempre sobre la version que hay ahora mismo en OneDrive */
     if(!forzar&&C.etag&&meta.eTag&&meta.eTag!==C.etag){
       var err=new Error("conflicto");err.conflicto=true;throw err;
     }
     C.etag=meta.eTag||C.etag;
     return descargarContenido();
   })
   .then(function(ab){
     ocupar("Cambiando las celdas…");
     return window.DINXW.aplicar(ab,plan);
   })
   .then(function(res){
     info=res;
     if(!res.escritas)throw new Error("no he encontrado ninguna celda que cambiar");
     ocupar("Subiendo el libro…");
     return subirContenido(res.bytes,forzar?null:C.etag);
   })
   .then(function(j){
     C.etag=(j&&j.eTag)||"";
     C.leido=new Date().toISOString();
     guardarCuenta();fijarBase();V.dirty=false;guardarLocalEstado();
     ocupar("");
     aviso("Escrito en el Excel: "+info.escritas+" celda"+(info.escritas===1?"":"s")+".");
     var pegas=(info.noEncontradas||[])
       .concat((info.hojasQueFaltan||[]).map(function(h){return "no existe la hoja «"+h+"»";}))
       .concat(p.sinSitio||[]);
     if(pegas.length)hoja("Esto se ha quedado fuera",
       '<p class="hint">El resto sí se ha escrito. Estas cosas hay que arreglarlas a mano en el Excel:</p>'+
       '<ul class="hint" style="margin-top:8px">'+
       pegas.map(function(s){return '<li style="margin:4px 0">· '+e(s)+'</li>';}).join("")+'</ul>');
   })
   .catch(function(err){
     ocupar("");
     if(err&&err.conflicto){
       hoja("El Excel ha cambiado por otro lado",
         '<p class="hint">El libro se ha guardado después de que la app lo leyera: puede haber sido tú '+
         'desde el ordenador, o quien tenga permiso de edición. Si escribo encima, se perderían esos cambios.</p>',
         '<button class="btn" data-act="releer">Traer el Excel y perder lo mío</button>'+
         '<button class="btn btn-danger" data-act="forzar">Escribir encima</button>'+
         '<button class="btn" data-act="cerrar">Dejarlo</button>');
       return;
     }
     aviso("No he podido guardar: "+((err&&err.message)||"error desconocido"));
   });
}


/* ---------------------------------------------------------------- descargas */
function bajar(nombre,datos,tipo){
  try{
    var b=new Blob([datos],{type:tipo||"text/plain;charset=utf-8"});
    var u=URL.createObjectURL(b),a=document.createElement("a");
    a.href=u;a.download=nombre;document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(u);a.remove();},400);
  }catch(x){aviso("No he podido preparar la descarga.");}
}

/* ---------------------------------------------------------------- tarjeta de cuenta */
function tarjetaCuenta(){
  var h='<div class="card card-pad" id="cuenta"><div class="sechead"><h2>Tu cuenta y tu Excel</h2>'+
    '<span class="right">'+(conectado()?e(C.cuenta||"conectado"):"sin conectar")+'</span></div>';

  if(!C.clientId){
    h+='<div class="note note-info"><span class="ic">1</span><div>Para que la app escriba en tu OneDrive '+
      'necesita un <b>ID de aplicación</b> de Microsoft. Se saca una vez, en cinco minutos: lo tienes '+
      'explicado paso a paso en el archivo <b>LEEME.md</b> que viene con la app.</div></div>'+
      '<div class="field" style="margin-top:14px"><label>ID de aplicación (Client ID)</label>'+
      '<input class="inp" data-act="clientid" placeholder="00000000-0000-0000-0000-000000000000" '+
      'spellcheck="false" autocomplete="off"></div>'+
      '<p class="hint-sm" style="margin-top:8px">No es una contraseña: las apps de navegador usan PKCE '+
      'y no llevan secreto. Se queda solo en este dispositivo.</p></div>';
    return h;
  }
  if(!conectado()){
    h+='<p class="hint" style="margin-bottom:14px">Entra con la cuenta de Microsoft donde tengas el libro.</p>'+
      '<button class="btn btn-primary" data-act="entrar">Entrar con Microsoft</button> '+
      '<button class="btn btn-quiet btn-sm" data-act="borrar-clientid">Cambiar el ID de aplicación</button></div>';
    return h;
  }
  if(!C.item){
    var migas='<button class="btn btn-sm" data-act="nav-raiz">Mi OneDrive</button>'+
      NAV.ruta.map(function(m,i){
        return ' <span class="hint-sm">/</span> <button class="btn btn-sm" data-act="nav-miga" data-i="'+i+'">'+
          e(m.nombre)+'</button>';
      }).join("");

    h+='<p class="hint" style="margin-bottom:12px">Dime cuál es tu libro. Puedes navegar por tus '+
      'carpetas, mirar los recientes o buscarlo por nombre.</p>'+
      '<div class="nav" style="flex-wrap:wrap;gap:6px;margin-bottom:12px">'+migas+
      ' <button class="btn btn-sm'+(NAV.modo==="recientes"?" btn-primary":"")+
      '" data-act="nav-recientes">Recientes</button></div>'+
      '<form class="quickadd" data-act="buscar" style="grid-template-columns:1fr auto;margin-bottom:4px">'+
      '<div class="field"><label>Buscar por nombre</label>'+
      '<input class="inp" name="q" value="Dineritos" autocomplete="off"></div>'+
      '<div class="field"><label>&nbsp;</label><button class="btn" type="submit">Buscar</button></div>'+
      '</form>';

    if(NAV.aviso)h+='<div class="note note-warn" style="margin-top:10px"><span class="ic">!</span><div>'+
      e(NAV.aviso)+'</div></div>';

    var filas="";
    NAV.carpetas.forEach(function(f){
      filas+='<tr><td class="c-name"><button class="btn btn-sm btn-quiet" data-act="nav-carpeta" '+
        'data-id="'+e(f.id)+'" data-nombre="'+e(f.name)+'">📁 '+e(f.name)+'</button></td>'+
        '<td class="hint-sm">carpeta</td><td></td></tr>';
    });
    NAV.archivos.forEach(function(f,i){
      var car=(f.parentReference&&f.parentReference.path||"")
        .replace(/^\/drive\/root:?\/?/,"").replace(/^\/drives\/[^/]+\/root:?\/?/,"");
      filas+='<tr><td class="c-name">'+e(f.name)+
        (car?'<small>'+e(car)+'</small>':'')+'</td>'+
        '<td class="hint-sm">'+e((f.lastModifiedDateTime||"").slice(0,10))+'</td>'+
        '<td><button class="btn btn-sm btn-primary" data-act="elegir" data-i="'+i+'">Es este</button></td></tr>';
    });
    if(filas)h+='<div class="tw" style="margin-top:12px"><table class="t"><tbody>'+filas+'</tbody></table></div>';
    else h+='<p class="hint" style="margin-top:12px">Aquí no hay carpetas ni hojas de cálculo.</p>';

    h+='<p class="hint-sm" style="margin-top:10px">Solo te enseño carpetas y archivos .xlsx / .xlsm; '+
      'el resto lo escondo para no estorbar.</p>'+
      '<hr class="divider"><button class="btn btn-quiet btn-sm" data-act="salir">Cerrar sesión</button></div>';
    return h;
  }

  var cuando=C.leido?new Date(C.leido).toLocaleString("es-ES"):"nunca";
  h+='<div class="tw"><table class="t"><tbody>'+
    '<tr><td class="c-name">Libro<small>en tu OneDrive</small></td><td class="n">'+e(C.nombre||"—")+'</td></tr>'+
    '<tr><td class="c-name">Última lectura</td><td class="n">'+e(cuando)+'</td></tr>'+
    '</tbody></table></div>'+
    '<div class="nav" style="flex-wrap:wrap;gap:8px;margin-top:14px">'+
    '<button class="btn btn-primary" data-act="guardar-excel">Guardar en el Excel</button>'+
    '<button class="btn" data-act="simular">Ver qué se va a escribir</button>'+
    '<button class="btn" data-act="releer">Leer del Excel</button>'+
    '<button class="btn btn-quiet btn-sm" data-act="otro-libro">Elegir otro libro</button>'+
    '<button class="btn btn-quiet btn-sm" data-act="salir">Cerrar sesión</button>'+
    '</div>'+
    '<div class="note note-info" style="margin-top:14px"><span class="ic">i</span><div>'+
    '<b>Guardar</b> baja el libro, cambia solo las celdas que toca y lo vuelve a subir. Todo lo demás '+
    'se copia byte a byte, así que tus gráficos, fórmulas y formatos quedan intactos, y Excel '+
    'recalcula los totales al abrirlo. OneDrive guarda la versión anterior en su historial. '+
    '<b>Leer</b> trae lo que diga el Excel y descarta lo que tengas aquí sin guardar.<br>'+
    'Cierra el libro en el ordenador antes de guardar, o Excel se quejará de coautoría.</div></div>'+
    '<hr class="divider">'+
    '<div class="sechead"><h2 style="font-size:15px">Sacar los datos</h2></div>'+
    '<div class="nav" style="flex-wrap:wrap;gap:8px">'+
    '<button class="btn btn-sm" data-act="exp-csv-apuntes">Apuntes en CSV</button>'+
    '<button class="btn btn-sm" data-act="exp-csv-resumen">Resumen del año en CSV</button>'+
    '<button class="btn btn-sm" data-act="exp-json">Copia de seguridad (JSON)</button>'+
    '</div></div>';
  return h;
}

/* Lo que este navegador no puede hacer, dicho antes de que falle algo */
function problemasNavegador(){
  var p=[];
  if(!window.isSecureContext||!(window.crypto&&crypto.subtle))
    p.push("La página no está en https, así que no puedo iniciar sesión con Microsoft.");
  if(typeof DecompressionStream==="undefined")
    p.push("Este navegador no sabe abrir archivos .xlsx. Necesitas Safari 16.4 o posterior, "+
           "o Chrome actualizado.");
  try{localStorage.setItem("din.test","1");localStorage.removeItem("din.test");}
  catch(x){p.push("No puedo guardar nada en este navegador: ¿estás en modo privado?");}
  return p;
}

/* ---------------------------------------------------------------- avisos */
var tT=null;
function aviso(msg){
  var v=document.querySelector(".toast");if(v)v.remove();
  var d=document.createElement("div");d.className="toast";d.textContent=msg;
  document.body.appendChild(d);
  clearTimeout(tT);tT=setTimeout(function(){d.remove();},4600);
}
function hoja(titulo,html,acciones){
  cerrarHoja();
  var s=document.createElement("div");s.className="scrim";s.id="scrim";
  s.innerHTML='<div class="sheet"><h3>'+e(titulo)+'</h3>'+html+
    '<div class="sheet-actions">'+(acciones||'<button class="btn" data-act="cerrar">Cerrar</button>')+'</div></div>';
  document.body.appendChild(s);
  s.addEventListener("click",function(ev){if(ev.target===s)cerrarHoja();});
}
function cerrarHoja(){var s=document.getElementById("scrim");if(s)s.remove();}
function ocupar(txt){OCUPADO=txt;pintar();}

/* ---------------------------------------------------------------- mutaciones */
function mutar(fn){
  fn();D.invalidar();V.dirty=true;guardarLocalEstado();pintar();
}
function mesActual(){
  var m=D.mesObj(V.mes,V.anio);
  if(!m){var ms=D.mesesOrden();m=ms[0];if(m){V.mes=m.mes;V.anio=m.anio;}}
  return m;
}
function numOrNull(v){
  if(v===""||v===null||v===undefined)return null;
  var n=parseFloat(String(v).replace(",","."));
  return isFinite(n)?n:null;
}

/* ---------------------------------------------------------------- estructura */
var TABS=[["mes","Mes"],["anio","Año"],["reparto","Reparto"],["anuales","Anuales"],["ajustes","Ajustes"]];
function nav(clase){
  return '<nav class="nav '+clase+'">'+TABS.map(function(t){
    return '<button data-act="tab" data-tab="'+t[0]+'"'+(V.tab===t[0]?' aria-current="page"':'')+'>'+t[1]+'</button>';
  }).join("")+'</nav>';
}
function estadoSync(){
  if(OCUPADO)return '<span class="saveflag"><i class="dot warn"></i>'+e(OCUPADO.replace("…",""))+'</span>';
  if(!listo())return '<button class="btn btn-sm" data-act="tab" data-tab="ajustes">Conectar</button>';
  if(V.dirty)return '<button class="btn btn-primary btn-sm" data-act="guardar-excel">Guardar</button>';
  return '<span class="saveflag"><i class="dot on"></i>al día</span>';
}
function pintar(){
  var raiz=document.getElementById("root"),cuerpo;
  /* Ajustes tiene que ser accesible ANTES de tener datos: es donde se pone el
     Id. de aplicacion y se conecta la cuenta. */
  var arranque=(!D.hayDatos()&&V.tab==="ajustes");
  if(arranque)cuerpo='<button class="btn btn-sm btn-quiet" data-act="tab" data-tab="mes"'+
    ' style="margin-top:16px">‹ Volver</button>'+U2.vistaAjustes();
  else if(!D.hayDatos())cuerpo=vistaInicio();
  else if(V.tab==="anio")cuerpo=U.vistaAnio();
  else if(V.tab==="reparto")cuerpo=U2.vistaReparto();
  else if(V.tab==="anuales")cuerpo=U2.vistaAnuales();
  else if(V.tab==="ajustes")cuerpo=U2.vistaAjustes();
  else cuerpo=U.vistaMes();

  var aviso2="";
  if(V.dirty&&listo())aviso2='<div class="note note-warn" style="margin:14px 0"><span class="ic">!</span>'+
    '<div>Tienes cambios que aún no están en el Excel. Pulsa <b>Guardar</b>.</div></div>';
  else if(D.hayDatos()&&!listo())aviso2='<div class="note note-info" style="margin:14px 0"><span class="ic">i</span>'+
    '<div>Trabajando solo en este dispositivo. Conecta tu cuenta en Ajustes para que escriba en el Excel.</div></div>';

  raiz.innerHTML=
    '<header class="topbar"><div class="topbar-in">'+
      '<div class="brand"><b>Dineritos</b><span>'+(D.hayDatos()?D.S.meses.length+" meses":"finanzas")+'</span></div>'+
      (D.hayDatos()?nav("navmob"):"")+estadoSync()+
    '</div></header>'+
    '<div class="shell">'+(D.hayDatos()?nav("navdesk"):"")+aviso2+cuerpo+
      '<footer class="foot">Tus datos van de este dispositivo a tu Excel de OneDrive y de vuelta. '+
      'Nada pasa por ningún otro servidor.'+
      '<br>Versión <b>'+VERSION+'</b>'+
      (C.leido?'<br>Última sincronización: '+new Date(C.leido).toLocaleString("es-ES"):'')+
      '</footer></div>'+
    (D.hayDatos()&&V.tab==="mes"?'<button class="fab" data-act="fab">＋ Anotar gasto</button>':'')+
    '<div class="tip" id="tip"></div>';
}
function vistaInicio(){
  var probs=problemasNavegador();
  return '<div class="mespick"><h1>Dineritos</h1></div>'+
    (probs.length?'<div class="note note-bad" style="margin-bottom:16px"><span class="ic">!</span><div>'+
      probs.map(function(s){return e(s);}).join("<br>")+'</div></div>':'')+
    '<div class="card card-pad" style="max-width:620px"><p class="hint" style="margin-bottom:16px">'+
    'Lleva tus finanzas con el mismo modelo que tu Excel y escribe en él directamente: '+
    'el reparto proporcional de los gastos comunes, el ahorro que se recalcula para llegar al objetivo del año, '+
    'los gastos anuales y la hucha.</p>'+
    (OCUPADO?'<p class="hint">'+e(OCUPADO)+'</p>':
      '<p class="hint" style="margin-bottom:14px">Empieza conectando tu cuenta de Microsoft.</p>'+
      '<button class="btn btn-primary" data-act="tab" data-tab="ajustes">Empezar</button>')+
    '<hr class="divider"><p class="hint-sm">También puedes trabajar sin cuenta: arrastra aquí el .xlsx '+
    'y quedará solo en este dispositivo.</p>'+
    '<div class="dz" data-act="dz" style="margin-top:10px"><b>Arrastra el .xlsx</b>'+
    '<span class="hint">o <button class="btn btn-sm" data-act="pick" style="vertical-align:baseline">elige el archivo</button></span>'+
    '<input type="file" id="fpick" accept=".xlsx" style="display:none"></div>'+
    '</div>';
}

/* ---------------------------------------------------------------- importar suelto */
function importarLocal(file){
  if(!file)return;
  if(!/\.xlsx$/i.test(file.name)){aviso("Eso no es un .xlsx.");return;}
  ocupar("Leyendo "+file.name+"…");
  D.abrirXlsx(file).then(function(st){
    D.S=st;guardarLocalEstado();fijarBase();
    V.dirty=false;V.mes=null;V.anio=null;V.anioVista=null;V.tab="mes";
    ocupar("");
    aviso("Importados "+st.meses.length+" meses.");
  },function(err){
    ocupar("");
    aviso((err&&err.message)||"No he podido leer el archivo.");
  });
}

window.DINPWA={
  pintar:pintar,aviso:aviso,hoja:hoja,cerrarHoja:cerrarHoja,mutar:mutar,mesActual:mesActual,
  numOrNull:numOrNull,bajar:bajar,tarjetaCuenta:tarjetaCuenta,cargarLocal:cargarLocal,
  guardarCuenta:guardarCuenta,guardarLocalEstado:guardarLocalEstado,fijarBase:fijarBase,
  entrar:entrar,salir:salir,procesarVuelta:procesarVuelta,buscarLibro:buscarLibro,
  leerExcel:leerExcel,guardarEnExcel:guardarEnExcel,importarLocal:importarLocal,
  ocupar:ocupar,pendientes:pendientes,conectado:conectado,listo:listo,
  cuenta:C,nav:NAV,verCarpeta:verCarpeta,subirNivel:subirNivel,verRecientes:verRecientes,
  simular:simular,planEscritura:planEscritura
};
})();


/* ===== p7_eventos.js ===== */
/* ==========================================================================
   Eventos y arranque del PWA
   ========================================================================== */
(function(){
"use strict";
var D=window.DIN,V=D.V,P=window.DINPWA;
var cap=D.cap,n0=D.n0,e=D.esc,eur=D.eur;
var pintar=P.pintar,aviso=P.aviso,hoja=P.hoja,cerrarHoja=P.cerrarHoja;
var mutar=P.mutar,mesActual=P.mesActual,numOrNull=P.numOrNull,bajar=P.bajar;

function onClick(ev){
  var t=ev.target.closest?ev.target.closest("[data-act]"):null;
  if(!t)return;
  var a=t.getAttribute("data-act"),i=+t.getAttribute("data-i"),m,ms;
  switch(a){
    /* ---- navegacion ---- */
    case "tab":
      V.tab=t.getAttribute("data-tab");pintar();window.scrollTo(0,0);break;
    case "mes-prev": case "mes-next":
      ms=D.mesesOrden();var ix=0;
      ms.forEach(function(x,k){if(x.mes===V.mes&&+x.anio===+V.anio)ix=k;});
      ix+=(a==="mes-next"?1:-1);
      if(ix>=0&&ix<ms.length){V.mes=ms[ix].mes;V.anio=ms[ix].anio;pintar();}
      break;
    case "ir-mes":
      ev.preventDefault();
      V.mes=t.getAttribute("data-mes");V.anio=+t.getAttribute("data-anio");V.tab="mes";
      pintar();window.scrollTo(0,0);break;
    case "tabla-cat": V.tablaCat=!V.tablaCat;pintar();break;
    /* desglose de una categoria: en Mes por medidor, en Ano por categoria */
    case "desglose":
      var dk=t.getAttribute("data-k");V.desglose=(V.desglose===dk?null:dk);pintar();break;
    case "desglose-anio":
      var dc=t.getAttribute("data-cat");V.desgloseAnio=(V.desgloseAnio===dc?null:dc);pintar();break;
    case "anio-modo": V.anioModo=t.getAttribute("data-modo");pintar();break;
    case "cerrar": cerrarHoja();break;
    case "fab":
      var f=document.querySelector('form[data-act="add-apunte"]');
      if(f){f.scrollIntoView({block:"center",behavior:"smooth"});
        var ic=f.querySelector('input[name=concepto]');
        if(ic)setTimeout(function(){ic.focus();},320);}
      break;

    /* ---- mes ---- */
    case "auto":
      var id=t.getAttribute("data-id");
      mutar(function(){
        m=mesActual();if(!m)return;
        m.reales=m.reales||{};
        if(m.reales[id]==="auto"){
          var fp=D.filasPlan(m.mes,m.anio).filter(function(x){return x.id===id;})[0];
          m.reales[id]=fp?Math.round(fp.plan*100)/100:0;
        }else m.reales[id]="auto";
      });break;
    case "del-apunte":
      mutar(function(){m=mesActual();if(m)m.apuntes.splice(i,1);});break;

    /* ---- listas editables ---- */
    case "rep-add":
      mutar(function(){D.S.reparto.push({grupo:"NUEVO",subconcepto:"",importe:0,notas:"",
        linkAnual:null,pagaDavidDirecto:false});});break;
    case "rep-del": mutar(function(){D.S.reparto.splice(i,1);});break;
    case "ga-add":
      mutar(function(){D.S.gastosAnuales.push({concepto:"NUEVO",importe:0,mes:"ENERO",ambito:"PERSONAL"});});break;
    case "ga-del":
      mutar(function(){
        var g=D.S.gastosAnuales[i];
        if(g)D.S.reparto.forEach(function(l){if(l.linkAnual===g.concepto)l.linkAnual=null;});
        D.S.gastosAnuales.splice(i,1);
      });break;
    case "rec-add":
      mutar(function(){D.S.recurrentes.push({concepto:"NUEVO",categoria:"PREVISTO",importe:0});});break;
    case "rec-del": mutar(function(){D.S.recurrentes.splice(i,1);});break;
    case "exc-add":
      mutar(function(){
        var ys=D.anios();
        D.S.excepciones.push({mes:"ENERO",anio:ys[ys.length-1]||new Date().getFullYear(),
          neto:n0(D.S.config.sueldoDavid),hacienda:n0(D.S.config.haciendaDefault)});
      });break;
    case "exc-del": mutar(function(){D.S.excepciones.splice(i,1);});break;
    case "curso-mes":
      mutar(function(){
        var y=t.getAttribute("data-anio"),mm=t.getAttribute("data-mes");
        D.S.mesesCurso=D.S.mesesCurso||{};
        var lst=D.S.mesesCurso[y]=D.S.mesesCurso[y]||[];
        var k=lst.indexOf(mm);
        if(k>=0)lst.splice(k,1);else lst.push(mm);
      });break;
    case "mes-del":
      var dm=t.getAttribute("data-mes"),da=+t.getAttribute("data-anio");
      mutar(function(){
        D.S.meses=D.S.meses.filter(function(x){return !(x.mes===dm&&+x.anio===da);});
        if(V.mes===dm&&+V.anio===da){V.mes=null;V.anio=null;}
      });break;

    /* ---- cuenta y Excel ---- */
    case "entrar": P.entrar();break;
    case "salir": P.salir();break;
    case "borrar-clientid":
      P.cuenta.clientId="";P.guardarCuenta();pintar();break;
    case "otro-libro":
      P.cuenta.item="";P.cuenta.drive="";P.cuenta.nombre="";P.cuenta.etag="";
      P.guardarCuenta();
      P.ocupar("Abriendo tu OneDrive…");
      P.verCarpeta(null).then(function(){P.ocupar("");},function(err){
        P.ocupar("");aviso("No he podido listar tu OneDrive: "+((err&&err.message)||""));});
      break;
    case "elegir":
      var f2=(P.nav.archivos||[])[i];
      if(!f2)break;
      P.cuenta.drive=(f2.parentReference&&f2.parentReference.driveId)||"";
      P.cuenta.item=f2.id;P.cuenta.nombre=f2.name;P.cuenta.etag="";
      P.guardarCuenta();
      if(!P.cuenta.drive){aviso("Ese resultado no trae identificador de unidad; prueba con otra búsqueda.");pintar();break;}
      P.ocupar("Leyendo el Excel…");
      P.leerExcel().then(function(st){
        P.ocupar("");V.tab="mes";
        aviso("Leídos "+st.meses.length+" meses de "+st.origen+".");
      },function(err){
        P.ocupar("");aviso("No he podido leerlo: "+((err&&err.message)||""));
      });
      break;
    case "nav-raiz":
      P.ocupar("Abriendo…");
      P.verCarpeta(null).then(function(){P.ocupar("");},function(err){
        P.ocupar("");aviso("No he podido abrir la carpeta: "+((err&&err.message)||""));});
      break;
    case "nav-carpeta":
      P.ocupar("Abriendo…");
      P.verCarpeta(t.getAttribute("data-id"),t.getAttribute("data-nombre"))
       .then(function(){P.ocupar("");},function(err){
         P.ocupar("");aviso("No he podido abrir la carpeta: "+((err&&err.message)||""));});
      break;
    case "nav-miga":
      P.ocupar("Abriendo…");
      P.subirNivel(i).then(function(){P.ocupar("");},function(err){
        P.ocupar("");aviso("No he podido abrir la carpeta: "+((err&&err.message)||""));});
      break;
    case "nav-recientes":
      P.ocupar("Buscando los recientes…");
      P.verRecientes().then(function(){P.ocupar("");},function(err){
        P.ocupar("");aviso("No he podido ver los recientes: "+((err&&err.message)||""));});
      break;
    case "simular": P.simular();break;
    case "guardar-excel": cerrarHoja();P.guardarEnExcel(false);break;
    case "forzar": cerrarHoja();P.guardarEnExcel(true);break;
    case "releer":
      cerrarHoja();
      if(V.dirty&&!confirm("Vas a traer lo que diga el Excel y se perderá lo que tengas aquí sin guardar. ¿Sigo?"))break;
      P.ocupar("Leyendo el Excel…");
      P.leerExcel().then(function(st){
        P.ocupar("");aviso("Al día: "+st.meses.length+" meses.");
      },function(err){
        P.ocupar("");aviso("No he podido leerlo: "+((err&&err.message)||""));
      });
      break;
    case "pick":
      var fp2=document.getElementById("fpick");if(fp2)fp2.click();break;

    /* ---- exportar ---- */
    case "exp-csv-apuntes":
      bajar("dineritos-apuntes.csv",D.csvApuntes(),"text/csv;charset=utf-8");break;
    case "exp-csv-resumen":
      var yv=V.anioVista||D.anios()[0];
      bajar("dineritos-resumen-"+yv+".csv",D.csvResumen(yv),"text/csv;charset=utf-8");break;
    case "exp-json":
      bajar("dineritos-copia.json",JSON.stringify(D.S,null,1),"application/json");break;
  }
}

function onChange(ev){
  var t=ev.target.closest?ev.target.closest("[data-act]"):null;
  if(!t)return;
  var a=t.getAttribute("data-act"),i=+t.getAttribute("data-i"),f=t.getAttribute("data-f");
  var v=t.value,m;
  switch(a){
    case "mes-sel":
      var ms=D.mesesOrden(),x=ms[+v];
      if(x){V.mes=x.mes;V.anio=x.anio;pintar();}break;
    case "anio-sel": V.anioVista=+v;pintar();break;
    case "clientid":
      var id=String(v||"").trim();
      if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)){
        if(id)aviso("Eso no parece un ID de aplicación: son 36 caracteres con guiones.");
        return;
      }
      P.cuenta.clientId=id;P.guardarCuenta();pintar();
      aviso("Guardado. Ya puedes entrar con Microsoft.");
      break;
    case "cfg": mutar(function(){D.S.config[t.getAttribute("data-k")]=n0(numOrNull(v));});break;
    case "cfg-mesbono": mutar(function(){D.S.config.mesBono=v||null;});break;
    case "cfg-aniobono": mutar(function(){D.S.config.anioBono=+v;});break;
    case "curso-imp":
      mutar(function(){
        D.S.config.cursoAleman=D.S.config.cursoAleman||{};
        D.S.config.cursoAleman[t.getAttribute("data-anio")]=n0(numOrNull(v));
      });break;
    case "real":
      mutar(function(){
        m=mesActual();if(!m)return;
        m.reales=m.reales||{};
        var nv=numOrNull(v);
        m.reales[t.getAttribute("data-id")]=nv===null?null:nv;
      });break;
    case "ap-edit":
      mutar(function(){
        m=mesActual();if(!m||!m.apuntes[i])return;
        m.apuntes[i][f]=(f==="importe")?n0(numOrNull(v)):(f==="fecha"?(v||null):v);
      });break;
    case "rep":
      mutar(function(){
        var l=D.S.reparto[i];if(!l)return;
        l[f]=(f==="importe")?n0(numOrNull(v)):v;
      });break;
    case "ga":
      mutar(function(){
        var g=D.S.gastosAnuales[i];if(!g)return;
        if(f==="concepto"){
          var antes=g.concepto;
          D.S.reparto.forEach(function(l){if(l.linkAnual===antes)l.linkAnual=v;});
        }
        g[f]=(f==="importe")?n0(numOrNull(v)):v;
      });break;
    case "rec":
      mutar(function(){
        var r=D.S.recurrentes[i];if(!r)return;
        r[f]=(f==="importe")?n0(numOrNull(v)):v;
      });break;
    case "exc":
      mutar(function(){
        var x2=D.S.excepciones[i];if(!x2)return;
        x2[f]=(f==="mes")?v:n0(numOrNull(v));
      });break;
  }
}

function onSubmit(ev){
  var t=ev.target.closest?ev.target.closest("[data-act]"):null;
  if(!t)return;
  var a=t.getAttribute("data-act");
  if(a==="add-apunte"){
    ev.preventDefault();
    var fd=new FormData(t);
    var concepto=String(fd.get("concepto")||"").trim();
    var imp=numOrNull(fd.get("importe"));
    if(!concepto||imp===null){aviso("Necesito al menos un concepto y un importe.");return;}
    mutar(function(){
      var m=mesActual();if(!m)return;
      m.apuntes=m.apuntes||[];
      m.apuntes.push({concepto:concepto,categoria:String(fd.get("categoria")||"VIDA DIARIA"),
        importe:imp,fecha:String(fd.get("fecha")||"")||null,notas:""});
    });
    aviso("Anotado: "+concepto+" · "+eur(imp));
    var f2=document.querySelector('form[data-act="add-apunte"]');
    if(f2){f2.reset();var ic=f2.querySelector('input[name=concepto]');if(ic)ic.focus();}
  }else if(a==="mes-add"){
    ev.preventDefault();
    var fd2=new FormData(t),mm=String(fd2.get("mes")),an=parseInt(fd2.get("anio"),10);
    if(!an){aviso("Pon un año válido.");return;}
    if(D.mesObj(mm,an)){aviso(cap(mm)+" "+an+" ya está en el plan.");return;}
    mutar(function(){D.S.meses.push({mes:mm,anio:an,reales:{},apuntes:[]});});
    aviso("Añadido "+cap(mm)+" "+an+". Ojo: no tendrá hoja en el Excel hasta que la crees allí.");
  }else if(a==="buscar"){
    ev.preventDefault();
    var q=String(new FormData(t).get("q")||"").trim();
    if(!q)return;
    P.ocupar("Buscando…");
    P.buscarLibro(q).then(function(){
      P.ocupar("");
    },function(err){
      P.ocupar("");aviso("La búsqueda ha fallado: "+((err&&err.message)||""));
    });
  }
}

/* ---------------------------------------------------------------- tooltips */
var tipEl=null;
function moverTip(ev){
  var t=ev.target.closest?ev.target.closest("[data-tip]"):null;
  if(!tipEl)tipEl=document.getElementById("tip");
  if(!tipEl)return;
  if(!t){tipEl.classList.remove("show");return;}
  tipEl.textContent=t.getAttribute("data-tip");
  tipEl.style.left=ev.clientX+"px";
  tipEl.style.top=ev.clientY+"px";
  tipEl.classList.add("show");
}
function salirTip(){if(tipEl)tipEl.classList.remove("show");}

/* ---------------------------------------------------------------- arranque */
function arrancar(){
  P.cargarLocal();
  pintar();

  document.addEventListener("click",onClick);
  document.addEventListener("change",onChange);
  document.addEventListener("submit",onSubmit);
  document.addEventListener("mousemove",moverTip,{passive:true});
  document.addEventListener("mouseleave",salirTip);
  document.addEventListener("scroll",salirTip,{passive:true});
  document.addEventListener("keydown",function(ev){
    if(ev.key==="Escape")cerrarHoja();
    if((ev.ctrlKey||ev.metaKey)&&ev.key==="s"){ev.preventDefault();if(V.dirty)P.guardarEnExcel(false);}
  });
  document.addEventListener("change",function(ev){
    if(ev.target&&ev.target.id==="fpick"&&ev.target.files&&ev.target.files[0])
      P.importarLocal(ev.target.files[0]);
  });
  ["dragenter","dragover"].forEach(function(tp){
    document.addEventListener(tp,function(ev){
      var dz=ev.target.closest?ev.target.closest('[data-act="dz"]'):null;
      if(dz){ev.preventDefault();dz.classList.add("hot");}
    });
  });
  document.addEventListener("dragleave",function(ev){
    var dz=ev.target.closest?ev.target.closest('[data-act="dz"]'):null;
    if(dz)dz.classList.remove("hot");
  });
  document.addEventListener("drop",function(ev){
    var dz=ev.target.closest?ev.target.closest('[data-act="dz"]'):null;
    if(!dz)return;
    ev.preventDefault();dz.classList.remove("hot");
    if(ev.dataTransfer&&ev.dataTransfer.files&&ev.dataTransfer.files[0])
      P.importarLocal(ev.dataTransfer.files[0]);
  });
  window.addEventListener("beforeunload",function(ev){
    if(V.dirty&&P.listo()){ev.preventDefault();ev.returnValue="";}
  });
  window.addEventListener("online",function(){
    if(V.dirty)aviso("Ya hay conexión: puedes guardar en el Excel.");
  });

  /* vuelta del login de Microsoft */
  P.procesarVuelta().then(function(res){
    if(res&&res.error){aviso("El login no ha salido: "+res.error);pintar();return;}
    if(res&&res.ok){
      V.tab="ajustes";
      if(P.listo()&&!D.hayDatos()){
        P.ocupar("Leyendo el Excel…");
        return P.leerExcel().then(function(st){
          P.ocupar("");V.tab="mes";
          aviso("Conectado. Leídos "+st.meses.length+" meses.");
        },function(err){
          P.ocupar("");aviso("Conectado, pero no he podido leer el libro: "+((err&&err.message)||""));
        });
      }
      aviso("Conectado"+(P.cuenta.cuenta?" como "+P.cuenta.cuenta:"")+".");
      if(!P.cuenta.item){
        P.ocupar("Abriendo tu OneDrive…");
        return P.verCarpeta(null).then(function(){P.ocupar("");},function(){P.ocupar("");});
      }
      pintar();
      return;
    }
    /* arranque normal: si hay libro enlazado y no hay datos, leerlo */
    if(P.listo()&&!D.hayDatos()){
      P.ocupar("Leyendo el Excel…");
      P.leerExcel().then(function(){P.ocupar("");},function(err){
        P.ocupar("");aviso("No he podido leer el libro: "+((err&&err.message)||""));
      });
    }
  });

  /* service worker: para que funcione sin conexion */
  if("serviceWorker" in navigator&&location.protocol==="https:"){
    navigator.serviceWorker.register("sw.js").catch(function(){});
  }
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",arrancar);
else arrancar();
})();

