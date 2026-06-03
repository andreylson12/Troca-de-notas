(async()=>{

const load=s=>new Promise(r=>{
  const x=document.createElement('script');
  x.src=s;
  x.onload=r;
  document.body.appendChild(x);
});

if(typeof XLSX==='undefined') await load('https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js');
if(typeof pdfjsLib==='undefined') await load('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js');
if(typeof Tesseract==='undefined') await load('https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js');

pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

const ROBO={xml:null,planilha:[],ordem:null,arquivos:{}};

const esperar=ms=>new Promise(r=>setTimeout(r,ms));

function normalizar(t){
  return String(t||'')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/[.\-\/,;:()]/g,' ')
    .replace(/\s+/g,' ')
    .trim();

}

function criarBotao(txt,top,cor){
  const b=document.createElement('button');
  b.innerText=txt;
  b.style=`position:fixed;top:${top}px;right:20px;z-index:999999;background:${cor};color:white;padding:12px 16px;border:0;border-radius:8px;font-weight:bold;cursor:pointer`;
  document.body.appendChild(b);
  return b;
}

function setInputCampo(el,valor){
  if(!el)return;
  el.focus();
  el.value='';
  el.dispatchEvent(new Event('input',{bubbles:true}));
  el.value=valor||'';
  el.dispatchEvent(new Event('input',{bubbles:true}));
  el.dispatchEvent(new Event('change',{bubbles:true}));
}

function setInput(id,valor){setInputCampo(document.querySelector('#'+id),valor);}

function setSelectIndex(id,index){
  const el=document.querySelector('#'+id);
  if(!el)return;
  el.selectedIndex=index;
  el.dispatchEvent(new Event('change',{bubbles:true}));
}

function setSelectTexto(id,texto){
  const el=document.querySelector('#'+id);
  if(!el)return false;
  const alvo=normalizar(texto);
  const opts=[...el.options];
  const achou=opts.findIndex(o=>normalizar(o.text).includes(alvo)||alvo.includes(normalizar(o.text)));
  if(achou>=0){
    el.selectedIndex=achou;
    el.dispatchEvent(new Event('change',{bubbles:true}));
    return true;
  }
  return false;
}

function escolherArquivo(accept){
  return new Promise(resolve=>{
    const input=document.createElement('input');
    input.type='file';
    input.accept=accept;
    input.onchange=()=>resolve(input.files[0]);
    input.click();
  });
}

function escolherPacote(){
  return new Promise(resolve=>{
    const input=document.createElement('input');
    input.type='file';
    input.multiple=true;
    input.accept='.xml,.xlsx,.xls,.xlsm,.csv,.pdf,image/*';
    input.onchange=()=>resolve([...input.files]);
    input.click();
  });
}

function nomeUF(sigla){
  const mapa={AC:'Acre',AL:'Alagoas',AP:'Amapá',AM:'Amazonas',BA:'Bahia',CE:'Ceará',DF:'Distrito Federal',ES:'Espírito Santo',GO:'Goiás',MA:'Maranhão',MT:'Mato Grosso',MS:'Mato Grosso do Sul',MG:'Minas Gerais',PA:'Pará',PB:'Paraíba',PR:'Paraná',PE:'Pernambuco',PI:'Piauí',RJ:'Rio de Janeiro',RN:'Rio Grande do Norte',RS:'Rio Grande do Sul',RO:'Rondônia',RR:'Roraima',SC:'Santa Catarina',SP:'São Paulo',SE:'Sergipe',TO:'Tocantins'};
  return mapa[String(sigla||'').toUpperCase()]||'';
}

function ajustarTipoVeiculo(tipo){
  return String(tipo||'')
  .replace(/RODOTREM\s*9\s*EIXOS?/i,'RODO-TREM 9 EIXO')
  .replace(/BITREM\s*7\s*EIXOS?/i,'BI-TREM 7 EIXO')
  .replace(/CARRETA\s*LS\s*6\s*EIXOS?/i,'CARRETA LS 6 EIXO')
  .trim();
}

function numeroLimpo(v){return String(v||'').replace(/[^\d]/g,'');}

async function arquivoParaCanvas(file){
  if(file.type.includes('image')){
    const img=new Image();
    img.src=URL.createObjectURL(file);
    await new Promise(r=>img.onload=r);
    const canvas=document.createElement('canvas');
    const ctx=canvas.getContext('2d');
    canvas.width=img.width;
    canvas.height=img.height;
    ctx.drawImage(img,0,0);
    return canvas;
  }

  const buffer=await file.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({data:buffer}).promise;
  const page=await pdf.getPage(1);
  const viewport=page.getViewport({scale:8});
  const canvas=document.createElement('canvas');
  const ctx=canvas.getContext('2d');
  canvas.width=viewport.width;
  canvas.height=viewport.height;
  await page.render({canvasContext:ctx,viewport}).promise;
  return canvas;
}

async function textoPDF(file){
  try{
    const buffer=await file.arrayBuffer();
    const pdf=await pdfjsLib.getDocument({data:buffer}).promise;
    const page=await pdf.getPage(1);
    const content=await page.getTextContent();
    return content.items.map(i=>i.str).join(' ');
  }catch(e){
    return '';
  }
}

async function ocrArquivo(file){
  const canvas=await arquivoParaCanvas(file);

  const c=document.createElement('canvas');
  const ctx=c.getContext('2d');

  c.width=canvas.width;
  c.height=canvas.height;
  ctx.drawImage(canvas,0,0);

  const result=await Tesseract.recognize(c,'por');

  console.log('OCR RESULTADO:');
  console.log(result.data.text);

  return result.data.text;
}

async function ocrPesagemFatal(file){
  const canvasOriginal=await arquivoParaCanvas(file);

  function girarCanvas(canvas,graus){
    const c=document.createElement('canvas');
    const ctx=c.getContext('2d');

    if(graus===90 || graus===270){
      c.width=canvas.height;
      c.height=canvas.width;
    }else{
      c.width=canvas.width;
      c.height=canvas.height;
    }

    ctx.translate(c.width/2,c.height/2);
    ctx.rotate(graus*Math.PI/180);
    ctx.drawImage(canvas,-canvas.width/2,-canvas.height/2);

    return c;
  }

  async function lerCanvas(canvas){
    const result=await Tesseract.recognize(canvas,'por');
    return result.data.text||'';
  }

  const tentativas=[
    await lerCanvas(canvasOriginal),
    await lerCanvas(girarCanvas(canvasOriginal,90)),
    await lerCanvas(girarCanvas(canvasOriginal,180)),
    await lerCanvas(girarCanvas(canvasOriginal,270))
  ];

  const melhor=tentativas
    .map(t=>({
      texto:t,
      pontos:
        (normalizar(t).includes('PESO')?5:0)+
        (normalizar(t).includes('INICIAL')?5:0)+
        (normalizar(t).includes('FINAL')?5:0)+
        (normalizar(t).includes('LIQUIDO')?5:0)+
        ((t.match(/\d{4,6}/g)||[]).length)
    }))
    .sort((a,b)=>b.pontos-a.pontos)[0];

  console.log('OCR PESAGEM FATAL:');
  console.log(melhor.texto);

  return melhor.texto;

}

async function carregarPacote(){
  const files=await escolherPacote();

  ROBO.arquivos={};

  for(const f of files){
    const nome=f.name.toLowerCase();

    if(nome.endsWith('.xml')){
      ROBO.arquivos.xml=f;
    }else if(nome.endsWith('.xlsx')||nome.endsWith('.xls')||nome.endsWith('.xlsm')||nome.endsWith('.csv')){
      ROBO.arquivos.planilha=f;
    }
  }

  const docs=files.filter(f=>!ROBO.arquivos.xml && false || (
    !f.name.toLowerCase().endsWith('.xml') &&
    !f.name.toLowerCase().endsWith('.xlsx') &&
    !f.name.toLowerCase().endsWith('.xls') &&
    !f.name.toLowerCase().endsWith('.xlsm') &&
    !f.name.toLowerCase().endsWith('.csv')
  ));

  for(const f of docs){
    let texto=await textoPDF(f);
    if(!texto || texto.length<50) texto=await ocrArquivo(f);

    const t=normalizar(texto);

    if(t.includes('PESAGEM') || t.includes('PESO LIQUIDO')){
      ROBO.arquivos.pesagem=f;
    }else if(t.includes('UMIDADE') || t.includes('IMPUREZAS') || t.includes('CLASSIFICACAO')){
      ROBO.arquivos.laudo=f;
    }else if(t.includes('ORDEM') || t.includes('CAVALO') || t.includes('MOTORISTA')){
      ROBO.arquivos.ordem=f;
    }
  }

  alert(
    'PACOTE CARREGADO:\n\n'+
    'XML: '+(ROBO.arquivos.xml?'OK':'FALTOU')+'\n'+
    'Planilha: '+(ROBO.arquivos.planilha?'OK':'FALTOU')+'\n'+
    'Ordem: '+(ROBO.arquivos.ordem?'OK':'FALTOU')+'\n'+
    'Laudo: '+(ROBO.arquivos.laudo?'OK':'FALTOU')+'\n'+
    'Pesagem: '+(ROBO.arquivos.pesagem?'OK':'FALTOU')
  );
}

async function lerXML(file=null){
  file=file||ROBO.arquivos.xml||await escolherArquivo('.xml');
  const txt=await file.text();
  const xml=new DOMParser().parseFromString(txt,'text/xml');
  const pegar=tag=>{
    const e=xml.getElementsByTagName(tag)[0];
    return e?e.textContent.trim():'';
  };

  ROBO.xml={
    chave:pegar('chNFe'),
    nf:pegar('nNF'),
    serie:pegar('serie'),
    produtor:pegar('xNome'),
    cpf:pegar('CPF'),
    fazenda:pegar('xFant'),
    peso:pegar('qCom'),
    placa:pegar('placa'),
    uf:pegar('UF')
  };

  console.log('XML:',ROBO.xml);
  alert('XML carregado:\n'+ROBO.xml.produtor);
}

async function lerPlanilha(file=null){
  file=file||ROBO.arquivos.planilha||await escolherArquivo('.xlsx,.xls,.xlsm,.csv');

  const buffer=await file.arrayBuffer();
  const wb=XLSX.read(buffer,{type:'array'});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const linhas=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});

  const dados=[];

  for(const r of linhas){
    const produtor=String(r[1]||'').trim();
    const bp=String(r[2]||'').trim();

    if(!produtor) continue;
    if(normalizar(produtor).includes('PRODUTOR')) continue;
    if(normalizar(produtor).includes('RESULTADOS')) continue;

    dados.push({
      centro:r[0]||'',
      produtor:r[1]||'',
      bp:r[2]||'',
      endRemessa:r[3]||'',
      bpRemessa:r[4]||'',
      op:r[5]||'',
      produtorSaida:r[6]||'',
      oflLv:r[7]||'',
      contrato:r[8]||'',
      opSaida:r[9]||'',
      descarga:r[10]||''
    });
  }

  ROBO.planilha=dados.filter(x=>x.produtor);

  console.log('PLANILHA CORRIGIDA:',ROBO.planilha);

  alert('Planilha carregada: '+ROBO.planilha.length+' registros válidos');
}

async function lerOrdem(file=null){
  file=file||ROBO.arquivos.ordem||await escolherArquivo('.pdf,image/*');

  let texto=await textoPDF(file);
if(!texto || texto.length<80) texto=await ocrArquivo(file);

console.log("========== TEXTO BRUTO ==========");
console.log(texto);
console.log("=================================");

  const textoLimpo=String(texto||'').replace(/\s+/g,' ').trim();

  function achar(...regexes){
    for(const rx of regexes){
      const m=textoLimpo.match(rx);
      if(m){
        for(let i=1;i<m.length;i++){
          if(m[i]) return String(m[i]).replace(/\s+/g,' ').trim();
        }
        return String(m[0]).replace(/\s+/g,' ').trim();
      }
    }
    return '';
  }

  function limparPlaca(v){
    return String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'').trim();
  }

  function somenteNumero(v){
    return String(v||'').replace(/[^\d]/g,'');
  }

  const ufs='AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO';

  const placasValidas=[...textoLimpo.matchAll(/[A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2}/g)]
    .map(m=>limparPlaca(m[0]))
    .filter(p=>p.length===7);

 const ehMotz=/MOTZ|HEBROM/i.test(textoLimpo);
const ehRodoviva=/RODOVIVA/i.test(textoLimpo);
const ehPampa=/RHPAMPA|PAMPA|MAISFRETE/i.test(textoLimpo);
const ehMafro=/MAFRO/i.test(textoLimpo);
const ehFribom=/FRIBON|FRIBOM/i.test(textoLimpo);

let placaCavalo='';
let placaCarreta1='';
let placaCarreta2='';
let placaCarreta3='';
let motorista='';
let cpfMotorista='';
let cnh='';
let uf='';
let tipoBruto='';
let tipoVeiculo='';
let transportadora='';

if(ehMotz){

  console.log('===== TEXTO MOTZ =====');
  console.log(textoLimpo);
  console.log('===== FIM TEXTO MOTZ =====');

  transportadora='MOTZ TRANSPORTES LTDA';

  const cpfMotzMatch=textoLimpo.match(/\d{3}\.\d{3}\.\d{3}\-\d{2}/);
  cpfMotorista=cpfMotzMatch ? cpfMotzMatch[0].replace(/\D/g,'') : '';

  if(cpfMotzMatch){
    const antesCpf=textoLimpo.slice(0, cpfMotzMatch.index).trim();
    const depoisCpf=textoLimpo.slice(cpfMotzMatch.index + cpfMotzMatch[0].length).trim();

    const motzNome=antesCpf.match(/TRANSPORTES\s+LTDA\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+)$/i);
    motorista=motzNome ? motzNome[1].trim() : '';

    const fimTrecho=depoisCpf.search(/GRANEL|PEDIDO|SOJA|EMITENTE|DATA|CNH|VOLUME/i);
    const trechoPlacas=fimTrecho>=0 ? depoisCpf.slice(0,fimTrecho) : depoisCpf;

    const placasMotz=[...trechoPlacas.matchAll(/[A-Z]{3}[-]?\d[A-Z0-9]\d{2}/g)]
      .map(x=>limparPlaca(x[0]))
      .filter(Boolean);

    placaCavalo=placasMotz[0]||'';
    placaCarreta1=placasMotz[1]||'';
    placaCarreta2=placasMotz[2]||'';
    placaCarreta3=placasMotz[3]||'';
  }

  // AJUSTE MOTZ MODELO NOVO:
  // procura placas no bloco PLACA CAVALO / PLACA CARRETA / PESO BRUTO
  const blocoMotzPlacas=textoLimpo.match(/PLACA\s+CAVALO[\s\S]{0,160}?PESO\s+BRUTO/i);

  if(blocoMotzPlacas){
    const placasBloco=[...blocoMotzPlacas[0].matchAll(/[A-Z]{3}\d[A-Z0-9]\d{2}/g)]
      .map(x=>limparPlaca(x[0]));

    if(placasBloco.length){
      placaCavalo=placasBloco[0]||placaCavalo;
      placaCarreta1=placasBloco[1]||placaCarreta1;
      placaCarreta2=placasBloco[2]||placaCarreta2;
      placaCarreta3=placasBloco[3]||placaCarreta3;
    }
  }

  // FALLBACK MOTZ MODELO COM PLACA CAVALO / PLACA CARRETA EM LINHAS
  if(!placaCavalo){
    placaCavalo=limparPlaca(
      achar(/PLACA\s+CAVALO\s*[:.\s]*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i)
    );
  }

  const carretasMotzLinha=[...textoLimpo.matchAll(/PLACA\s+CARRETA\s*[:.\s]*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/gi)]
    .map(x=>limparPlaca(x[1]))
    .filter(Boolean);

  if(carretasMotzLinha.length){
    placaCarreta1=placaCarreta1||carretasMotzLinha[0]||'';
    placaCarreta2=placaCarreta2||carretasMotzLinha[1]||'';
    placaCarreta3=placaCarreta3||carretasMotzLinha[2]||'';
  }

  if(!motorista){
    motorista=achar(
      /CLIENTE\s*:\s*AGREX\s+D[EO]\s+BRASIL\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+SOJA/i,
      /AGREX\s+D[EO]\s+BRASIL\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+SOJA/i,
      /MOTORISTA\s*[:.\s]*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i,
      /MOTORISTA\s*[:.\s]*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+\d{3}[.\s]?\d{3}[.\s]?\d{3}/i
    );
  }

  if(!cpfMotorista){
    cpfMotorista=somenteNumero(
      achar(/CPF\s*[:.\s]*([\d\.\/\-]{9,20})/i)
    );
  }

  const ufsMotz=[...textoLimpo.matchAll(/\bUF\s*:?\s*(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b/gi)];
  uf=ufsMotz.length ? ufsMotz[ufsMotz.length-1][1].toUpperCase() : '';

  if(!uf){
    uf=achar(/FILIAL\s*:\s*([A-Z]{2})/i).toUpperCase();
  }

  if(!uf){
    const ufsLinha=[...textoLimpo.matchAll(/CIDADE\s*:\s*(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)/gi)];
    uf=ufsLinha.length ? ufsLinha[ufsLinha.length-1][1].toUpperCase() : '';
  }

  cnh=somenteNumero(
    achar(/CNH\s*[:.\s]*(\d{5,15})/i)
  );

  tipoBruto='RODOTREM 9 EIXO';
  tipoVeiculo='RODO-TREM 9 EIXO';

}
else if(ehFribom){

  console.log('===== TEXTO FRIBOM =====');
console.log(textoLimpo);
console.log('===== FIM TEXTO =====');

  transportadora='FRIBON TRANSPORTES LTDA PIAUÍ';

motorista=achar(
  /Proprietario\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+QUANTIDADE/i,
  /Proprietario\s*:\s*[A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?\s+MONTE\s+ALEGRE-PI\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+VOLUMES/i
);

const docFribom = textoLimpo.match(/(\d{11})\s+(\d{5,10})\s+(\d{8,15})/);

if(docFribom){
  cpfMotorista=docFribom[1];
  cnh=docFribom[3];
}
  if(!motorista){
  motorista=achar(
    /MOTORISTA\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i,
    /ORDEM\s+DE\s+CARREGAMENTO\s+AGREX\s+DO\s+BRASIL\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ\s\-]+?\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ\-]+\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+\d{9,11}[-]?\d{0,2}\s+\d{8,15}/i
  );
}

const docFribom2 = textoLimpo.match(/([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+(\d{9,11})[-]?(\d{0,2})\s+(\d{8,15})\s+RODOTREM/i);

if(docFribom2){
  if(!motorista) motorista=docFribom2[1].trim();
  cpfMotorista=(docFribom2[2]+docFribom2[3]).replace(/\D/g,'');
  cnh=docFribom2[4].replace(/\D/g,'');
}
  placaCavalo=limparPlaca(
    achar(/ve[ií]culo\s+placa\s*[:.\s]*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i)
  );

  if(!placaCavalo){
    placaCavalo=placasValidas[0]||'';
  }

  const carretas=[...new Set(
  [...textoLimpo.matchAll(/PFK\d[A-Z0-9]{2}/gi)]
    .map(x=>limparPlaca(x[0]))
)];

placaCarreta1=carretas[0]||'';
placaCarreta2=carretas[1]||'';
placaCarreta3=carretas[2]||'';

  uf=achar(
    /Estado\s*[:.\s]*([A-Z]{2})/i,
    /\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b\s+Carreta/i
  ).toUpperCase();

  if(/BI[\-\s]*TREM\s*7\s*EIXO/i.test(textoLimpo)){
    tipoVeiculo='BI-TREM 7 EIXO';
    tipoBruto='BI-TREM 7 EIXO';
  }else if(/RODO[\-\s]*TREM\s*9\s*EIXO/i.test(textoLimpo)){
    tipoVeiculo='RODO-TREM 9 EIXO';
    tipoBruto='RODO-TREM 9 EIXO';
  }else{
    tipoVeiculo='CARRETA LS 6 EIXO';
    tipoBruto='CARRETA LS 6 EIXO';
  }

}

else if(ehMafro){

  transportadora='MAFRO TRANSPORTES LTDA';

 motorista=achar(
  /MAFRO TRANSPORTES LTDA\s+\(PI\)\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+RIA/i,
  /Dados do Motorista[\s\S]{0,120}?([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{3,}(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}){2,})\s+[A-Z]{3}[-]?\d[A-Z0-9]\d{2}/i,
  /Motorista\s*[:.\s]*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i
);

  cpfMotorista=somenteNumero(
    achar(/CPF\s*[:.\s]*([\d\.\/\-]+)/i)
  );

  cnh=somenteNumero(
    achar(/CNH\s*[:.\s]*(\d{5,15})/i)
  );

  let placaTemp=limparPlaca(
    achar(/Placa\s*\(cavalo\)\s*[:.\s]*([A-Z0-9\-]+)/i)
  );

  if(!/^[A-Z]{3}\d[A-Z0-9]\d{2}$/.test(placaTemp)){
    placaTemp='';
  }

  placaCavalo = placaTemp || (placasValidas.find(p=>p!=='UF') || '');

  const carretasTodas=[...textoLimpo.matchAll(/[A-Z]{3}\d[A-Z0-9]\d{2}/g)]
    .map(x=>limparPlaca(x[0]))
    .filter(p=>p!==placaCavalo);

  placaCarreta1=carretasTodas[0]||'';
  placaCarreta2=carretasTodas[1]||'';
  placaCarreta3=carretasTodas[2]||'';

  uf=achar(
    /Placa\s*\(cavalo\).*?UF\s*[:.\s]*([A-Z]{2})/i,
    new RegExp('UF\\s*[:.\\s]*('+ufs+')','i')
  ).toUpperCase();

  if(/RODOTREM\s*9\s*EIXOS?|RODO\s*TREM\s*9\s*EIXOS?|9EIXOS/i.test(textoLimpo)){
    tipoVeiculo='RODO-TREM 9 EIXO';
    tipoBruto='RODOTREM 9 EIXOS';
  }else if(/BITREM|BI\s*TREM|7\s*EIXOS?/i.test(textoLimpo)){
    tipoVeiculo='BI-TREM 7 EIXO';
    tipoBruto='BI-TREM 7 EIXO';
  }else{
    tipoVeiculo='CARRETA LS 6 EIXO';
    tipoBruto='CARRETA LS 6 EIXO';
  }

}
else if(ehPampa){

  transportadora='RHPAMPA TRANSPORTES LTDA';

  placaCavalo=limparPlaca(
    achar(/CAVALO:\s*([A-Z0-9\-]+)/i)
  );

  placaCarreta1=limparPlaca(
    achar(/CARRETA\s*1:\s*([A-Z0-9\-]+)/i)
  );

  placaCarreta2=limparPlaca(
    achar(/CARRETA\s*2:\s*([A-Z0-9\-]+)/i)
  );

  motorista=achar(
    /MOTORISTA:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i
  );

  cpfMotorista=somenteNumero(
    achar(/CPF:\s*([\d\.\/\-]+)/i)
  );

  cnh=somenteNumero(
    achar(/N[°º]?\s*CNH\s*DOC\.?\s*:\s*(\d+)/i)
  );

  uf=achar(
    /CAVALO:.*?\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\s*\//i
  ).toUpperCase();

  const eixos=parseInt(
    achar(/TOTAL\s*EIXOS:\s*(\d+)/i)
  ) || 0;

  if(eixos>=9){
    tipoBruto='RODOTREM 9 EIXO';
    tipoVeiculo='RODO-TREM 9 EIXO';
  }
  else if(eixos===7){
    tipoBruto='BITREM 7 EIXO';
    tipoVeiculo='BI-TREM 7 EIXO';
  }
  else{
    tipoBruto='CARRETA LS';
    tipoVeiculo='CARRETA LS 6 EIXO';
  }

}

else if(ehRodoviva){

  transportadora='RODOVIVA TRANSPORTES LTDA';

  placaCavalo=limparPlaca(achar(
    /Cavalo\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
  ));

  placaCarreta1=limparPlaca(achar(
    /Carreta\s*1\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
  ));

  placaCarreta2=limparPlaca(achar(
    /Carreta\s*2\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
  ));

  motorista=achar(
    /Solicitamos\s+entregar\s+ao\s+motorista\s+Sr\.?\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i,
    /motorista\s+Sr\.?\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i
  );

  cpfMotorista=somenteNumero(achar(
    /CPF\s*:\s*([\d\.\-\/]+)/i
  ));

  cnh=somenteNumero(achar(
    /CNH\s*:\s*(\d{5,15})/i
  ));

  uf=achar(
    new RegExp('UF\\s*:\\s*('+ufs+')','i')
  ).toUpperCase();

if(/RODO\s*TREM\s*9|RODOTREM\s*9/i.test(textoLimpo)){
    tipoBruto='RODO TREM 9 EIXO';
    tipoVeiculo='RODO-TREM 9 EIXO';
  }

} // FECHA RODOVIVA

else{

  transportadora=achar(
      /(TRANSPORTES\s+FOB\s+LTDA)/i,
      /(FRIBOM[^\s]*\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ\s\.]*)/i,
      /(MOTZ[^\s]*\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ\s\.]*)/i,
      /Embarcador\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s\.]+?)\s+CNPJ/i
  );
    placaCavalo=limparPlaca(achar(
      /Placa\s*cavalo\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i,
      /Cavalo\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
    ));

    placaCarreta1=limparPlaca(achar(
      /Placa\s*carreta\s*1\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i,
      /Carreta\s*1\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
    ));

    placaCarreta2=limparPlaca(achar(
      /Placa\s*carreta\s*2\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i,
      /Carreta\s*2\s*:\s*([A-Z]{3}[-\s]?\d[A-Z0-9][-\s]?\d{2})/i
    ));

    motorista=achar(
      /Motorista\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+(?:Endere[cç]o|CPF|RG|Bairro|CEP|Cidade|Fone|Org[aã]o|Dt\.?\s*Emi)/i,
      /Motorista\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+(?:Endere[cç]o|CPF|RG|Bairro|CEP|Cidade|Fone|Org[aã]o|Dt\.?\s*Emi)/i
    );

    cpfMotorista=somenteNumero(
      ([...textoLimpo.matchAll(/CPF\s*:\s*([\d\.\-\/]{9,20})/gi)]
        .map(m=>somenteNumero(m[1]))
        .filter(n=>n.length===11)
        .pop()) || ''
    );

    cnh=somenteNumero(achar(
      /Form\.?\s*C\.?\s*N\.?\s*H\.?\s*[:\-]?\s*(\d{5,15})/i,
      /C\.?\s*N\.?\s*H\.?\s*[:\-]?\s*(\d{5,15})/i,
      /CNH\s*[:\-]?\s*(\d{5,15})/i
    ));

    uf=achar(
      new RegExp('Cidade\\s*:\\s*[A-ZÁÉÍÓÚÂÊÔÃÕÇ\\s]+?\\s+UF\\s*[:\\-]?\\s*('+ufs+')','i'),
      new RegExp('UF\\s*[:\\-]?\\s*('+ufs+')','i')
    ).toUpperCase();

    tipoBruto=achar(
      /Obs\.?\s*1\s*:\s*([A-Z0-9ÁÉÍÓÚÂÊÔÃÕÇ\s\-]+?)(?:Obs\.?\s*2|$)/i,
      /(RODO[\-\s]*TREM\s*9\s*EIXOS?)/i,
      /(RODOTREM\s*9\s*EIXOS?)/i,
      /(BI[\-\s]*TREM\s*7\s*EIXOS?)/i,
      /(BITREM\s*7\s*EIXOS?)/i,
      /(LS\s*4\s*EIXOS?\s*CARRETA)/i,
      /(CARRETA\s*LS\s*6\s*EIXOS?)/i,
      /Tipo\s*:\s*([A-Z0-9ÁÉÍÓÚÂÊÔÃÕÇ\s\-]+?)\s+(?:No\.?\s*MCT|Renavam|Modelo|Ano|Placa|Obs|$)/i
    );

    const tipoN=normalizar(tipoBruto);

    if(tipoN.includes('RODO') || tipoN.includes('RODOTREM')){
      tipoVeiculo='RODO-TREM 9 EIXO';
    }else if(tipoN.includes('BITREM') || tipoN.includes('BI TREM')){
      tipoVeiculo='BI-TREM 7 EIXO';
    }else if(tipoN.includes('LS') || tipoN.includes('4 EIXOS CARRETA')){
      tipoVeiculo='CARRETA LS 6 EIXO';
    }else if(tipoN.includes('CAMINHAO TRATOR')){
      tipoVeiculo='CARRETA LS 6 EIXO';
    }else{
      tipoVeiculo=ajustarTipoVeiculo(tipoBruto);
    }
  }

  if(!placaCavalo && placasValidas.length) placaCavalo=placasValidas[0];
  if(!placaCarreta1 && placasValidas.length>1) placaCarreta1=placasValidas.find(p=>p!==placaCavalo)||'';
  if(!placaCarreta2 && placasValidas.length>2) placaCarreta2=placasValidas.find(p=>p!==placaCavalo && p!==placaCarreta1)||'';

  if(!uf && ROBO.xml && ROBO.xml.uf) uf=String(ROBO.xml.uf).toUpperCase();

  motorista=String(motorista||'')
    .replace(/\bCPF\b.*$/i,'')
    .replace(/\bENDERECO\b.*$/i,'')
    .replace(/\bENDEREÇO\b.*$/i,'')
    .replace(/\bBAIRRO\b.*$/i,'')
    .replace(/\bFONE\b.*$/i,'')
    .trim();

  if(!tipoVeiculo){
    if(/RODO\s*TREM\s*9|RODOTREM\s*9/i.test(textoLimpo)){
      tipoVeiculo='RODO-TREM 9 EIXO';
      tipoBruto='RODO TREM 9 EIXO';
    }
  }

  const mercadoria=achar(
    /Mercadoria\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s]+?)\s+Esp[eé]cie/i,
    /Mercadoria\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s]+?)\s+Quant/i,
    /PRODUTO\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s]+?)\s+CATEGORIA/i
  );

  const especie=achar(
    /Esp[eé]cie\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9\s]+?)\s+(?:Quant|Peso|Autorizamos|$)/i
  );

  ROBO.ordem={
    texto:textoLimpo,
    placasEncontradas:placasValidas,
    placa:placaCavalo,
    placaCavalo,
    placaCarreta:placaCarreta1,
    placaCarreta1,
    placaCarreta2,
    placaCarreta3,
    uf,
    ufNome:nomeUF(uf),
    motorista,
    cpfMotorista,
    cnh,
    tipoBruto,
    tipoVeiculo,
    mercadoria,
    especie,
    transportadora
  };

  console.log('ORDEM CORRIGIDA:',ROBO.ordem);

  alert(
    'Ordem carregada:\n\n'+
    'Transportadora: '+(ROBO.ordem.transportadora||'NÃO ACHOU')+'\n'+
    'Placa cavalo: '+(ROBO.ordem.placaCavalo||'NÃO ACHOU')+'\n'+
    'Carreta 1: '+(ROBO.ordem.placaCarreta1||'NÃO ACHOU')+'\n'+
    'Carreta 2: '+(ROBO.ordem.placaCarreta2||'NÃO ACHOU')+'\n'+
    'Carreta 3: '+(ROBO.ordem.placaCarreta3||'')+'\n'+
    'UF: '+(ROBO.ordem.uf||'NÃO ACHOU')+' '+(ROBO.ordem.ufNome||'')+'\n'+
    'Motorista: '+(ROBO.ordem.motorista||'NÃO ACHOU')+'\n'+
    'CPF: '+(ROBO.ordem.cpfMotorista||'NÃO ACHOU')+'\n'+
    'CNH: '+(ROBO.ordem.cnh||'NÃO ACHOU')+'\n'+
    'Tipo: '+(ROBO.ordem.tipoVeiculo||ROBO.ordem.tipoBruto||'NÃO ACHOU')
  );
}
function buscarRemessa(){
  if(!ROBO.xml)return null;

  const prod=normalizar(ROBO.xml.produtor)
    .replace(/\bS A\b/g,'SA')
    .replace(/\bLTDA\b/g,'LTDA');

  let linha=ROBO.planilha.find(l=>{
    const nome=normalizar(l.produtor)
      .replace(/\bS A\b/g,'SA')
      .replace(/\bLTDA\b/g,'LTDA');

    return nome.includes(prod) || prod.includes(nome);
  });

  if(!linha){
    linha=ROBO.planilha.find(l=>{
      const nome=normalizar(l.produtor)
        .replace(/\bS A\b/g,'SA')
        .replace(/\bLTDA\b/g,'LTDA');

      const partes=prod.split(' ').filter(p=>p.length>2);
      return partes.every(p=>nome.includes(p));
    });
  }

  console.log('REMESSA ENCONTRADA:',linha);
  return linha;
}
async function selecionarEnderecoRemessaPorBP(remessa){
  const codigoRemessa = String(remessa.bpRemessa || remessa.endRemessa || '').replace(/\D/g,'');

  if(!codigoRemessa){
    alert('BP Remessa vazio na planilha.');
    return false;
  }

  const campoPrincipal = document.querySelector('#enderecoRemessa');

  if(!campoPrincipal){
    alert('Não achei o campo principal Endereço de Remessa.');
    return false;
  }

  const lupa =
    campoPrincipal.parentElement.querySelector('.fa-search, i[class*="search"], button') ||
    campoPrincipal.closest('div').querySelector('.fa-search, i[class*="search"], button');

  if(!lupa){
    alert('Não achei a lupa do Endereço de Remessa.');
    return false;
  }

  lupa.click();
  await esperar(1500);

  const campoBusca = document.querySelector('#filtroBusca');

  if(!campoBusca){
    alert('Não achei o campo filtroBusca da janela.');
    return false;
  }

  setInputCampo(campoBusca,'');
  await esperar(300);

  setInputCampo(campoBusca,codigoRemessa);
  await esperar(800);

  campoBusca.dispatchEvent(new KeyboardEvent('keyup',{bubbles:true,key:'Enter',code:'Enter'}));
  campoBusca.dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'Enter',code:'Enter'}));
  campoBusca.dispatchEvent(new Event('change',{bubbles:true}));

  await esperar(2500);

  const btnSelecionar = [...document.querySelectorAll('button')]
    .find(b =>
      b.offsetParent !== null &&
      normalizar(b.innerText || '').includes('SELECIONAR')
    );

  if(!btnSelecionar){
    alert('Não achei o botão Selecionar depois de buscar BP '+codigoRemessa+'.');
    return false;
  }

  btnSelecionar.click();
  await esperar(1200);

  return true;
}

async function preencherPrimeiraTela(){
  const remessa=buscarRemessa();
  if(!remessa)return alert('Produtor não encontrado na planilha.');

  setSelectIndex('unidadeNegocio',3); await esperar(500);
  setSelectIndex('processo',1); await esperar(500);
  setSelectIndex('tipoTransporte',4); await esperar(500);

  const placaFinal=(ROBO.ordem&&ROBO.ordem.placaCavalo)||ROBO.ordem.placa||ROBO.xml.placa;
  setInput('placa',placaFinal); await esperar(500);

  const ufFinal=ROBO.ordem.ufNome||nomeUF(ROBO.xml.uf);
  if(ufFinal)setSelectTexto('uf',ufFinal);
  await esperar(500);

  if(!setSelectTexto('tipoVeiculo',ROBO.ordem.tipoVeiculo)){
    if(ROBO.ordem.tipoVeiculo.includes('RODO'))setSelectTexto('tipoVeiculo','RODO-TREM 9 EIXO');
    else if(ROBO.ordem.tipoVeiculo.includes('BI'))setSelectTexto('tipoVeiculo','BI-TREM 7 EIXO');
    else if(ROBO.ordem.tipoVeiculo.includes('CARRETA'))setSelectTexto('tipoVeiculo','CARRETA LS 6 EIXO');
    else setSelectIndex('tipoVeiculo',2);
  }
  await esperar(500);

  setInput('nomeMotorista',ROBO.ordem.motorista); await esperar(500);

  setInput('cpfMotorista',ROBO.ordem.cpfMotorista||''); await esperar(300);
  setInput('cnh',ROBO.ordem.cnh||''); await esperar(300);
  setInput('placaCavalo',ROBO.ordem.placaCavalo||ROBO.ordem.placa||''); await esperar(300);
  setInput('placaCarreta',ROBO.ordem.placaCarreta||''); await esperar(300);

  setSelectIndex('operacao',4); await esperar(500);
  setSelectIndex('material',3); await esperar(500);
  setSelectIndex('transgenia',2); await esperar(500);
  setSelectIndex('safra',6); await esperar(500);
  setSelectIndex('deposito',3); await esperar(500);

  const okRemessa = await selecionarEnderecoRemessaPorBP(remessa);

  if(!okRemessa){
    alert('Endereço de Remessa não selecionado. Corrija antes de gerar ticket.');
    return;
  }

  alert('Primeira tela preenchida com BP Remessa. Confira e clique em Gerar Ticket.');
}

async function preencherNF(){
  const remessa=buscarRemessa();
  const produtorCodigo=(remessa&&remessa.bp)?String(remessa.bp).replace(/\D/g,''):'';

  if(!ROBO.xml)return alert('XML não carregado.');
  if(!produtorCodigo)return alert('Não achei BP do produtor na planilha.');

  // 1) Clicar em Nova Nota Fiscal
  const btnNovaNF=[...document.querySelectorAll('button')]
    .find(b=>normalizar(b.innerText).includes('NOVA NOTA FISCAL'));

  if(btnNovaNF){
    btnNovaNF.click();
    await esperar(2500);
  }

  // 2) Selecionar produtor pelo BP
  let campoProdutor =
    document.querySelector('input[aria-autocomplete="list"]') ||
    document.querySelector('input[placeholder*="Produtor"]') ||
    document.querySelector('input[formcontrolname*="produtor"]');

  if(campoProdutor){
    setInputCampo(campoProdutor,'');
    await esperar(300);

    setInputCampo(campoProdutor,produtorCodigo);
    await esperar(2200);

    let opcoes=[...document.querySelectorAll('.ng-option, .ng-dropdown-panel .ng-option')]
      .filter(o=>o.offsetParent!==null);

    let opcao=opcoes.find(o=>
      normalizar(o.innerText).includes(normalizar(produtorCodigo))
    );

    if(!opcao && opcoes.length) opcao=opcoes[0];

    if(opcao){
      opcao.scrollIntoView({block:'center'});
      await esperar(300);
      opcao.click();
      await esperar(1500);
    }
  }

  // 3) Selecionar fazenda, se existir
  const fazenda=document.querySelector('#fazenda');

  if(fazenda && fazenda.options && fazenda.options.length>1){
    fazenda.selectedIndex=1;
    fazenda.dispatchEvent(new Event('change',{bubbles:true}));
    await esperar(1200);
  }

  // 4) Preencher chave da NF
  const campoChave =
    document.querySelector('input[formcontrolname="chave"]') ||
    document.querySelector('input[formcontrolname*="chave"]') ||
    document.querySelector('input[placeholder*="chave"]') ||
    [...document.querySelectorAll('input')]
      .find(i=>normalizar(i.placeholder||'').includes('CHAVE'));

  if(!campoChave){
    return alert('Não achei o campo da chave da NF.');
  }

  setInputCampo(campoChave,'');
  await esperar(300);

  setInputCampo(campoChave,ROBO.xml.chave);
  await esperar(1200);

  // 5) Clicar na lupa/buscar da chave
  const grupoChave =
    campoChave.closest('.input-group') ||
    campoChave.parentElement ||
    document;

  const lupaChave =
    grupoChave.querySelector('.fa-search, .fa-search-plus, i[class*="search"], button') ||
    [...document.querySelectorAll('.fa-search, .fa-search-plus, i[class*="search"], button')]
      .reverse()
      .find(e=>e.offsetParent!==null);

  if(lupaChave){
    lupaChave.click();
    await esperar(2500);
  }

  alert('NF preenchida. Confira os dados e salve a NF manualmente.');
}

function pegarNumero(texto,regex){
  const m=texto.match(regex);
  if(!m)return '0.00';
  return m[1].replace(',','.').trim();
}

async function lerLaudoClassificacao(file=null){
  file=file||ROBO.arquivos.laudo||await escolherArquivo('.pdf,image/*');

  alert('Lendo laudo... aguarde.');

let texto=await textoPDF(file);

console.log('TEXTO PDF:', texto);

if(!texto || texto.length<200){
  console.log('PDF sem texto. Usando OCR...');
  texto=await ocrArquivo(file);
}

  console.log('OCR LAUDO:',texto);

  const dados={
    TaxaTicket70:pegarNumero(texto,/Umidade\s*:\s*(\d+[,.]\d+)/i),
    TaxaTicket71:(()=>{
      const m=texto.match(/(?:Mat.*?Impurezas?|Impurezas?)\s*:\s*(\d+[,.]\d+)/i);
      if(!m)return '0.00';
      let bruto=m[1];
      if(bruto.includes('89,'))bruto=bruto.replace('89,','0,');
      let n=parseFloat(bruto.replace(',','.'));
      if(n>10)n=n/100;
      return n.toFixed(2);
    })(),
    TaxaTicket72:pegarNumero(texto,/Esverdeados\s*:\s*(\d+[,.]\d+)/i),
    TaxaTicket73:pegarNumero(texto,/Ardidos\s*:\s*(\d+[,.]\d+)/i),
    TaxaTicket74:pegarNumero(texto,/Quebrados\s*\/?\s*Amassados\s*:\s*(\d+[,.]\d+)/i),
    TaxaTicket75:pegarNumero(texto,/Total\s*(?:de\s*)?Avariad[oa]s?\s*:\s*(\d+[,.]\d+)/i)
  };

  const aba=[...document.querySelectorAll('*')]
    .find(e=>e.innerText&&e.innerText.trim()==='Classificação');

  if(aba){aba.click(); await esperar(1500);}

  for(const id in dados){
    setInputCampo(document.querySelector('#'+id),dados[id]);
  }

  alert('Classificação preenchida. Confira antes de salvar.');
}

function extrairPesos(texto){
  console.log('OCR PESAGEM:',texto);

  const raw=String(texto||'');

  function num(v){
    return parseInt(String(v||'').replace(/\D/g,''),10)||0;
  }

  function achar(...regexes){
    for(const rx of regexes){
      const m=raw.match(rx);
      if(m && m[1]){
        const n=num(m[1]);
        if(n>=1000 && n<=100000) return n;
      }
    }
    return 0;
  }

  function retorno(tara,bruto,liquido,metodo){
    return {
      tara:String(tara),
      bruto:String(bruto),
      liquido:String(liquido || (bruto-tara)),
      metodo
    };
  }

  // REGRA 0: validação matemática com qualquer trio de números
  const todos=[...raw.matchAll(/\b\d{2,3}[.,]\d{3}\b|\b\d{5}\b/g)]
    .map(m=>num(m[0]))
    .filter(v=>v>=10000 && v<=90000);

  const unicos=[...new Set(todos)].sort((a,b)=>a-b);

  console.log('PESOS CANDIDATOS:',unicos);

  for(const bruto of unicos.slice().reverse()){
    for(const tara of unicos){
      if(bruto<=tara) continue;
      const calc=bruto-tara;

      for(const liquido of unicos){
        if(Math.abs(calc-liquido)<=150){
          return retorno(tara,bruto,liquido,'REGRA 0 - TRIO VALIDADO');
        }
      }
    }
  }

  // REGRA 1: PESAGEM ENTRADA / PESAGEM SAIDA
  const entrada=achar(
    /PESAGEM\s+ENTRADA[\s\S]{0,120}?PESO\s*[:|]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /ENTRADA[\s\S]{0,120}?PESO\s*[:|]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  const saida=achar(
    /PESAGEM\s+SA[IÍ]DA[\s\S]{0,120}?PESO\s*[:|]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /SA[IÍ]DA[\s\S]{0,120}?PESO\s*[:|]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  const totalEntradaSaida=achar(
    /TOTAL\s+L[IÍ]QUIDO[\s\S]{0,80}?(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  if(entrada && saida && entrada>=10000 && saida>=10000){
    const tara=Math.min(entrada,saida);
    const bruto=Math.max(entrada,saida);
    const calc=bruto-tara;
    const liquido=(totalEntradaSaida && Math.abs(totalEntradaSaida-calc)<=500)
      ? totalEntradaSaida
      : calc;

    return retorno(tara,bruto,liquido,'REGRA 1 - ENTRADA/SAIDA');
  }

  // REGRA 2: PESO INICIAL / PESO FINAL
  const inicial=achar(
    /PESO\s+INICIAL\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /INICIAL\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  const final=achar(
    /PESO\s+FINAL\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /FINAL\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  const liquidoInicialFinal=achar(
    /PESO\s+L[IÍE]QUIDO\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /LIQUIDO\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i,
    /LEQUIDO\s*[:=]?\s*(\d{2,3}[.,]?\d{3}|\d{5})/i
  );

  if(inicial && final && inicial>=10000 && final>=10000){
    const tara=Math.min(inicial,final);
    const bruto=Math.max(inicial,final);
    const calc=bruto-tara;
    const liquido=(liquidoInicialFinal && Math.abs(liquidoInicialFinal-calc)<=500)
      ? liquidoInicialFinal
      : calc;

    return retorno(tara,bruto,liquido,'REGRA 2 - INICIAL/FINAL');
  }

  // REGRA 3: PESO BRUTO / PESO TARA / LIQUIDO
  const brutoRotulo=achar(
    /PESO\s+BRUTO\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i,
    /BRUTO\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i
  );

  const taraRotulo=achar(
    /PESO\s+TARA\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i,
    /TARA\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i
  );

  const liquidoRotulo=achar(
    /LIQUIDO\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i,
    /L[IÍ]QUIDO\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i,
    /SALDO\s*[:.\-=]*\s*(\d{2,3}[.,]\d{3}|\d{5})/i
  );

  if(brutoRotulo && taraRotulo){
    const calc=brutoRotulo-taraRotulo;
    const liquido=(liquidoRotulo && Math.abs(liquidoRotulo-calc)<=500)
      ? liquidoRotulo
      : calc;

    return retorno(taraRotulo,brutoRotulo,liquido,'REGRA 3 - BRUTO/TARA');
  }

  // REGRA 4: maior e menor apenas se tiver dois pesos confiáveis
  if(unicos.length>=2){
    const tara=unicos[0];
    const bruto=unicos[unicos.length-1];
    return retorno(tara,bruto,bruto-tara,'REGRA 4 - MAIOR/MENOR');
  }

  alert('Não consegui validar a pesagem automaticamente. Vou pedir manual.');

  return {
    tara:prompt('Digite a tara / menor peso:', ''),
    bruto:prompt('Digite o peso bruto / maior peso:', ''),
    liquido:'',
    metodo:'MANUAL'
  };

  }

  function processoEntrada(){
  const proc=document.querySelector('#processo');
  const texto=proc?proc.options[proc.selectedIndex].text:'';
  return texto.toUpperCase().includes('ENTRADA');
}
async function clicarGerarTicket(){
  await esperar(1200);

  const btnGerar=[...document.querySelectorAll('button')]
    .find(b =>
      b.offsetParent!==null &&
      normalizar(b.innerText).includes('GERAR TICKET')
    );

  if(!btnGerar){
    alert('Não achei o botão Gerar Ticket.');
    return false;
  }

  if(btnGerar.disabled){
    alert('Botão Gerar Ticket ainda está desabilitado.');
    return false;
  }

btnGerar.click();

console.log('Ticket gerado. Aguardando carregar...');

await esperar(8000);
}

async function lerPesagemOCR(file=null){
  file=file||ROBO.arquivos.pesagem||await escolherArquivo('.pdf,image/*');

  alert('Lendo pesagem... aguarde.');

  let texto=await textoPDF(file);

  if(!texto || texto.length<80){
    texto=await ocrPesagemFatal(file);
  }

  const pesos=extrairPesos(texto);

  alert(
    'Pesagem lida:\n\n'+
    'Bruto: '+pesos.bruto+'\n'+
    'Tara: '+pesos.tara+'\n'+
    'Líquido: '+(pesos.liquido||'')+'\n'+
    'Método: '+pesos.metodo
  );

  const aba=[...document.querySelectorAll('*')]
    .find(e=>e.innerText&&e.innerText.trim()==='Pesagem');

  if(aba){aba.click(); await esperar(1000);}

  const inputs=[...document.querySelectorAll('input')]
    .filter(i=>!i.disabled&&i.offsetParent!==null);

  const campo1=inputs.find(i=>(i.placeholder||'').toLowerCase().includes('primeiro'))||inputs[0];
  const campo2=inputs.find(i=>(i.placeholder||'').toLowerCase().includes('segundo'))||inputs[1];

  let primeiro,segundo;

  if(processoEntrada()){
    primeiro=pesos.bruto;
    segundo=pesos.tara;
  }else{
    primeiro=pesos.tara;
    segundo=pesos.bruto;
  }

  setInputCampo(campo1,primeiro); await esperar(500);
  setInputCampo(campo2,segundo); await esperar(500);

  const btn=[...document.querySelectorAll('button')]
    .find(b=>b.innerText.trim().includes('Pesar'));

  if(btn)btn.click();

  alert('Pesagem preenchida. Confira antes de salvar.');
}
async function continuarDepoisTicket(){
  if(confirm('Abrir/preencher NF agora?')){
    await preencherNF();
  }

  alert('Confira e SALVE A NF manualmente.\nDepois clique OK.');

  if(confirm('Preencher CLASSIFICAÇÃO agora?')){
    await lerLaudoClassificacao(ROBO.arquivos.laudo);
  }

  alert('Confira e salve a classificação manualmente.\nDepois clique OK.');

  if(confirm('Preencher PESAGEM agora?')){
    await lerPesagemOCR(ROBO.arquivos.pesagem);
  }

  alert('CONTINUAÇÃO FINALIZADA.\nConfira tudo antes de salvar/finalizar.');
}

async function executarGuiado(){
  if(!ROBO.arquivos.xml || !ROBO.arquivos.planilha || !ROBO.arquivos.ordem || !ROBO.arquivos.laudo || !ROBO.arquivos.pesagem){
    return alert('Primeiro clique em CARREGAR PACOTE e selecione os 5 arquivos.');
  }

  await lerXML(ROBO.arquivos.xml);
  await lerPlanilha(ROBO.arquivos.planilha);
  await lerOrdem(ROBO.arquivos.ordem);

  if(confirm('Preencher PRIMEIRA TELA e gerar ticket agora?')){
    await preencherPrimeiraTela();

    const gerouTicket=await clicarGerarTicket();

    if(!gerouTicket){
      alert('Primeira tela preenchida. Gere o ticket manualmente.');
      return;
    }

    alert('Ticket gerado. Quando a próxima tela abrir, clique em CONTINUAR.');
    return;
  }
}
 
function criarPainelRobo(){
  const box=document.createElement('div');
  box.style=`
    position:fixed;
    top:100px;
    right:20px;
    z-index:999999;
    background:white;
    border:1px solid #ccc;
    border-radius:10px;
    padding:10px;
    width:170px;
    box-shadow:0 4px 12px rgba(0,0,0,.2);
    font-family:Arial;
  `;

  box.innerHTML=`
    <div style="font-weight:bold;margin-bottom:8px;text-align:center;color:#064e3b">
      🤖 ROBÔ SOLUM
    </div>
  `;

  function btn(txt,cor,fn){
    const b=document.createElement('button');
    b.innerText=txt;
    b.style=`
      width:100%;
      margin:4px 0;
      padding:8px;
      border:0;
      border-radius:6px;
      background:${cor};
      color:white;
      font-weight:bold;
      cursor:pointer;
      font-size:12px;
    `;
    b.onclick=fn;
    box.appendChild(b);
  }

  btn('📦 Carregar Pacote','#9333ea',carregarPacote);
  btn('▶ Executar Guiado','#111827',executarGuiado);
  btn('⏭ Continuar','#2563eb',continuarDepoisTicket);

  const detalhes=document.createElement('details');
  detalhes.style='margin-top:6px;';
  detalhes.innerHTML='<summary style="cursor:pointer;font-weight:bold;">Etapas</summary>';

  const area=document.createElement('div');
  detalhes.appendChild(area);
  box.appendChild(detalhes);

  function btnEtapa(txt,cor,fn){
    const b=document.createElement('button');
    b.innerText=txt;
    b.style=`
      width:100%;
      margin:3px 0;
      padding:7px;
      border:0;
      border-radius:6px;
      background:${cor};
      color:white;
      font-weight:bold;
      cursor:pointer;
      font-size:11px;
    `;
    b.onclick=fn;
    area.appendChild(b);
  }

  btnEtapa('1 - XML','#065f46',()=>lerXML());
  btnEtapa('2 - PLANILHA','#7c3aed',()=>lerPlanilha());
  btnEtapa('3 - ORDEM','#ca8a04',()=>lerOrdem());
  btnEtapa('4 - 1ª TELA','#1d4ed8',preencherPrimeiraTela);
  btnEtapa('5 - NF','#0f766e',preencherNF);
  btnEtapa('6 - LAUDO','#b91c1c',()=>lerLaudoClassificacao());
  btnEtapa('7 - PESAGEM','#0ea5e9',()=>lerPesagemOCR());

  document.body.appendChild(box);
}
criarPainelRobo();

alert('ROBÔ SOLUM V2 CORRIGIDO - PAINEL COMPACTO CARREGADO.');
  })();
