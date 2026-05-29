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
  return String(t||'').toUpperCase().normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'').replace(/-/g,' ')
  .replace(/\s+/g,' ').trim();
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
  const result=await Tesseract.recognize(canvas,'por');
  return result.data.text;
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

   const nomeArq=f.name.toLowerCase();

if(
  nomeArq.includes('laudo') ||
  nomeArq.includes('classificacao') ||
  nomeArq.includes('classificação')
){
  ROBO.arquivos.laudo=f;
}
else if(
  nomeArq.includes('balanca') ||
  nomeArq.includes('balança') ||
  nomeArq.includes('pesagem') ||
  t.includes('PESAGEM') ||
  t.includes('PESO LIQUIDO')
){
  ROBO.arquivos.pesagem=f;
}
else if(
  nomeArq.includes('oc') ||
  nomeArq.includes('ordem') ||
  t.includes('ORDEM') ||
  t.includes('CAVALO') ||
  t.includes('MOTORISTA')
){
  ROBO.arquivos.ordem=f;
}
else if(
  t.includes('UMIDADE') ||
  t.includes('IMPUREZAS') ||
  t.includes('CLASSIFICACAO')
){
  ROBO.arquivos.laudo=f;
}
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
  for(let i=1;i<linhas.length;i++){
    const r=linhas[i];
    dados.push({
      centro:r[0],
      produtor:r[1],
      bp:r[2],
      endRemessa:r[3],
      bpRemessa:r[4],
      op:r[5]
    });
  }

  ROBO.planilha=dados.filter(x=>x.produtor);
  console.log('PLANILHA:',ROBO.planilha);
  alert('Planilha carregada: '+ROBO.planilha.length+' registros');
}

async function lerOrdem(file=null){
  file=file||ROBO.arquivos.ordem||await escolherArquivo('.pdf,image/*');

  let texto=await textoPDF(file);
  if(!texto || texto.length<50) texto=await ocrArquivo(file);

  let placa='';
 const mPlaca=
  texto.match(/CARRO\s+DE\s+PLACA\s*(?:N[º°]?\s*)?([A-Z]{3}\d[A-Z0-9]\d{2})/i) ||
  texto.match(/CAVALO\s*:\s*([A-Z]{3}\d[A-Z0-9]\d{2})/i) ||
  texto.match(/P\.\s*CAVALO\s*:\s*([A-Z]{3}\d[A-Z0-9]\d{2})/i) ||
  texto.match(/PLACA\s*DO\s*VE[IÍ]CULO[\s\S]{0,50}?([A-Z]{3}\d[A-Z0-9]\d{2})/i) ||
  texto.match(/[A-Z]{3}\d[A-Z0-9]\d{2}/);

  if(mPlaca) placa=mPlaca[1]||mPlaca[0];

  let uf='';
  const mUfCavalo=texto.match(/CAVALO\s*:\s*[A-Z]{3}\d[A-Z0-9]\d{2}[\s\S]{0,160}?\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\s*\/\s*[A-Z]/i);
  const mUfPlaca=texto.match(/PLACA\s*DO\s*VE[IÍ]CULO[\s\S]{0,100}?([A-Z]{3}\d[A-Z0-9]\d{2})[\s\S]{0,80}?\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b/i);

  if(mUfCavalo) uf=mUfCavalo[1].toUpperCase();
  else if(mUfPlaca) uf=mUfPlaca[2].toUpperCase();

  let motorista='';
 const mMotorista=
  texto.match(/MOTORISTA\s+SR\.?\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i) ||
  texto.match(/MOTORISTA\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i) ||
  texto.match(/Motorista\s*:\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+Contato/i) ||
  texto.match(/MOTORISTA\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+?)\s+CPF/i);

  if(mMotorista) motorista=mMotorista[1].trim();

  let tipo='';
  const mTipo=
    texto.match(/RODOTREM\s*9\s*EIXOS?/i) ||
    texto.match(/BITREM\s*7\s*EIXOS?/i) ||
    texto.match(/CARRETA\s*LS\s*6\s*EIXOS?/i) ||
    texto.match(/TIPO DE VEICULO\s*:\s*([A-Z0-9\s]+?)(?:Data|$)/i);

  if(mTipo) tipo=mTipo[1]?mTipo[1].trim():mTipo[0].trim();

  ROBO.ordem={
    texto,
    placa,
    uf,
    ufNome:nomeUF(uf),
    motorista,
    tipoVeiculo:ajustarTipoVeiculo(tipo)
  };

  console.log('ORDEM:',ROBO.ordem);
  alert('Ordem carregada:\nPlaca: '+ROBO.ordem.placa+'\nUF: '+ROBO.ordem.ufNome+'\nMotorista: '+ROBO.ordem.motorista+'\nTipo: '+ROBO.ordem.tipoVeiculo);
}

function buscarRemessa(){
  if(!ROBO.xml)return null;
  const prod=normalizar(ROBO.xml.produtor);

  let linha=ROBO.planilha.find(l=>{
    const nome=normalizar(l.produtor);
    return nome.includes(prod)||prod.includes(nome);
  });

  if(!linha){
    linha=ROBO.planilha.find(l=>{
      const nome=normalizar(l.produtor);
      return prod.split(' ').every(p=>nome.includes(p));
    });
  }

  console.log('REMESSA:',linha);
  return linha;
}

async function preencherPrimeiraTela(){
  const remessa=buscarRemessa();
  if(!remessa)return alert('Produtor não encontrado na planilha.');

  setSelectIndex('unidadeNegocio',3); await esperar(500);
  setSelectIndex('processo',1); await esperar(500);
  setSelectIndex('tipoTransporte',4); await esperar(500);

  const placaFinal=ROBO.xml.placa||ROBO.ordem.placa;
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
  setSelectIndex('operacao',4); await esperar(500);
  if(!setSelectTexto('material','SOJA EM GRÃOS')){
  if(!setSelectTexto('material','SOJA EM GRAOS')){
    if(!setSelectTexto('material','SOJA')){
      setSelectIndex('material',3);
    }
  }
}
await esperar(500);
  setSelectIndex('transgenia',2); await esperar(500);
  setSelectIndex('safra',6); await esperar(500);
  setSelectIndex('deposito',3); await esperar(500);

  setInput('enderecoRemessa',remessa.endRemessa||remessa.bpRemessa);
  await esperar(1500);

  const opcao=[...document.querySelectorAll('.ng-option')]
  .find(o=>normalizar(o.innerText).includes(normalizar(remessa.endRemessa))||
           normalizar(o.innerText).includes(normalizar(remessa.bpRemessa)));

  if(opcao)opcao.click();

  alert('Primeira tela preenchida. Confira e clique em Gerar Ticket.');
}

async function preencherNF(){
  const remessa=buscarRemessa();
  const produtorCodigo=(remessa&&remessa.bp)?String(remessa.bp):'';

  if(!produtorCodigo)return alert('Não achei BP do produtor na planilha.');

  const btnNovaNF=[...document.querySelectorAll('button')]
    .find(b=>b.innerText.trim()==='Nova Nota Fiscal');

  if(btnNovaNF){
    btnNovaNF.click();
    await esperar(2200);
  }

  const campoProdutor=document.querySelector('input[aria-autocomplete="list"]');

  if(campoProdutor){
    setInputCampo(campoProdutor,produtorCodigo);
    await esperar(1800);

    const opcao=[...document.querySelectorAll('.ng-option')]
      .find(o=>o.innerText.includes(produtorCodigo));

    if(opcao)opcao.click();
  }

  await esperar(1200);

  const fazenda=document.querySelector('#fazenda');
  if(fazenda){
    fazenda.selectedIndex=1;
    fazenda.dispatchEvent(new Event('change',{bubbles:true}));
  }

  await esperar(1200);

  const campoChave=
    document.querySelector('input[formcontrolname="chave"]')||
    document.querySelector('input[placeholder*="chave"]');

  setInputCampo(campoChave,ROBO.xml.chave);
  await esperar(1200);

  const lupa=document.querySelector('.fa-search-plus.fa-input');
  if(lupa)lupa.click();

  alert('NF preenchida pelo XML. Confira e salve a NF manualmente.');
}

function pegarNumero(texto,regex){
  const m=texto.match(regex);
  if(!m)return '0.00';
  return m[1].replace(',','.').trim();
}

async function lerLaudoClassificacao(file=null){
  file=file||ROBO.arquivos.laudo||await escolherArquivo('.pdf,image/*');

  alert('Lendo laudo... aguarde.');

  const texto=await ocrArquivo(file);

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

  const candidatos=[...texto.matchAll(/\b\d{2}[.,]\d{3}\b/g)]
    .map(m=>m[0])
    .map(n=>parseInt(numeroLimpo(n),10))
    .filter(v=>v>=20000&&v<=90000);

  const valores=[...new Set(candidatos)].sort((a,b)=>a-b);

  if(valores.length>=3){
    return {tara:String(valores[0]),bruto:String(valores[valores.length-1]),metodo:'3 pesos encontrados'};
  }

  if(valores.length===2){
    const menor=valores[0], maior=valores[1];
    return {tara:String(maior-menor),bruto:String(maior),metodo:'calculado: bruto - líquido'};
  }

  return {
    tara:prompt('Digite a tara / menor peso:', ''),
    bruto:prompt('Digite o peso bruto / maior peso:', ''),
    metodo:'manual'
  };
}

function processoEntrada(){
  const proc=document.querySelector('#processo');
  const texto=proc?proc.options[proc.selectedIndex].text:'';
  return texto.toUpperCase().includes('ENTRADA');
}

async function lerPesagemOCR(file=null){
  file=file||ROBO.arquivos.pesagem||await escolherArquivo('.pdf,image/*');

  alert('Lendo pesagem... aguarde.');

  const texto=await ocrArquivo(file);
  const pesos=extrairPesos(texto);

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

async function executarGuiado(){
  if(!ROBO.arquivos.xml || !ROBO.arquivos.planilha || !ROBO.arquivos.ordem || !ROBO.arquivos.laudo || !ROBO.arquivos.pesagem){
    return alert('Primeiro clique em 9 - CARREGAR PACOTE e selecione os 5 arquivos.');
  }

  await lerXML(ROBO.arquivos.xml);
  await lerPlanilha(ROBO.arquivos.planilha);
  await lerOrdem(ROBO.arquivos.ordem);

  if(confirm('Preencher PRIMEIRA TELA agora?')){
    await preencherPrimeiraTela();
  }

  alert('Confira a primeira tela e clique em GERAR TICKET manualmente.\nDepois clique OK.');

  if(confirm('Já gerou o ticket? Abrir/preencher NF agora?')){
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

  alert('EXECUÇÃO GUIADA FINALIZADA.\nConfira tudo antes de salvar/finalizar.');
}

criarBotao('1 - XML',120,'#065f46').onclick=()=>lerXML();
criarBotao('2 - PLANILHA',170,'#7c3aed').onclick=()=>lerPlanilha();
criarBotao('3 - ORDEM',220,'#ca8a04').onclick=()=>lerOrdem();
criarBotao('4 - PREENCHER 1ª TELA',270,'#1d4ed8').onclick=preencherPrimeiraTela;
criarBotao('5 - ABRIR/PREENCHER NF',320,'#0f766e').onclick=preencherNF;
criarBotao('6 - LAUDO / CLASSIFICAÇÃO',370,'#b91c1c').onclick=()=>lerLaudoClassificacao();
criarBotao('7 - PESAGEM OCR',420,'#0ea5e9').onclick=()=>lerPesagemOCR();
criarBotao('8 - EXECUTAR GUIADO',470,'#111827').onclick=executarGuiado;
criarBotao('9 - CARREGAR PACOTE',520,'#9333ea').onclick=carregarPacote;

alert('ROBÔ SOLUM V1 FINAL AJUSTADO CARREGADO.');

})();
