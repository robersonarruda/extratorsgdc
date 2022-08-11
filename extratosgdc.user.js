// ==UserScript==
// @name          Extrator Contatos Sigeduca
// @fullname      Extrator Contatos Sigeduca
// @version       2.1.2
// @description   Consulta e salva dados de contato dos alunos do sigeduca.
// @author        Roberson Arruda
// @homepage      https://github.com/robersonarruda/extratorsgdc/blob/main/extratosgdc.user.js
// @downloadURL   https://github.com/robersonarruda/extratorsgdc/raw/main/extratosgdc.user.js
// @updateURL     https://github.com/robersonarruda/extratorsgdc/raw/main/extratosgdc.user.js
// @include	      *sigeduca.seduc.mt.gov.br/ged/hwmconaluno.aspx*
// @copyright     2019, Roberson Arruda (robersonarruda@outlook.com)
// @grant         none
// ==/UserScript==


//CARREGA libJquery
var libJquery = document.createElement('script');
libJquery.src = 'https://code.jquery.com/jquery-3.4.0.min.js';
libJquery.language='javascript';
libJquery.type = 'text/javascript';
document.getElementsByTagName('head')[0].appendChild(libJquery);

//CSS DOS BOTÕES
var styleSCT = document.createElement('style');
styleSCT.type = 'text/css';
styleSCT.innerHTML =
'.botaoSCT {'+
'	-moz-box-shadow:inset 1px 1px 0px 0px #b2ced4;'+
'	-webkit-box-shadow:inset 1px 1px 0px 0px #b2ced4;'+
'	box-shadow:inset 1px 1px 0px 0px #b2ced4;'+
'	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #4e88ed), color-stop(1, #3255c7) );'+
'	background:-moz-linear-gradient( center top, #4e88ed 5%, #3255c7 100% );'+
'	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#4e88ed", endColorstr="#3255c7");'+
'	background-color:#4e88ed;'+
'	-moz-border-radius:4px;'+
'	-webkit-border-radius:4px;'+
'	border-radius:4px;'+
'	border:1px solid #102b4d;'+
'	display:inline-block;'+
'	color:#ffffff;'+
'	font-family:Trebuchet MS;'+
'	font-size:11px;'+
'	font-weight:bold;'+
'	padding:2px 0px;'+
'	width:152px;'+
'	text-decoration:none;'+
'	text-shadow:1px 1px 0px #100d29;'+
'}.botaoSCT:hover {'+
'	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #3255c7), color-stop(1, #4e88ed) );'+
'	background:-moz-linear-gradient( center top, #3255c7 5%, #4e88ed 100% );'+
'	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#3255c7", endColorstr="#4e88ed");'+
'	background-color:#3255c7;'+
'}.botaoSCT:active {'+
'	position:relative;'+
'	top:1px;}'+
'.menuSCT{'+
'	-moz-border-radius:4px;'+
'	-webkit-border-radius:4px;'+
'	border-radius:4px;'+
'	border:1px solid #102b4d;}'
document.getElementsByTagName('head')[0].appendChild(styleSCT);

//Variáveis
var vetAluno = [0];
var n = 0;
var a = "";
var k = 0;
var cabecalho;

//FUNÇÃO SALVAR CONTEÚDO EM TEXTO
function saveTextAsFile() {
    var conteudo = document.getElementById("txtDados").value; //Retira acentos =>> .normalize('NFD').replace(/[\u0300-\u036f]/g, "");
    var a = document.createElement('a');
    with (a) {
        var href='data:text/csv;base64,' + btoa(conteudo);
        var download= 'dadosGED.csv';
    }
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function destroyClickedElement(event) {
  document.body.removeChild(event.target);
}


function coletar(opcao)
{
    n=0;
    vetAluno = [0];
    vetAluno = txtareaAluno.value;
    vetAluno = vetAluno.replace(/\n/g, ",");
    vetAluno = vetAluno.replace(/\t/g, ",");
    vetAluno = vetAluno.replace(/ /g,",");
    vetAluno = vetAluno.replace(/;/g, ",");
    vetAluno = vetAluno.replace(/\./g, ",");
    vetAluno = vetAluno.replace(/,,/g, ",");
    vetAluno = vetAluno.replace(/,,/g, ",");
    vetAluno = vetAluno.replace(/,,/g, ",");
    vetAluno = vetAluno.split(",");
    k=vetAluno.length;
    while(k >= 0){
        if(vetAluno[k]==""){vetAluno.splice(k, 1)}
            k--;
        if(k < 0){
            a = "";
            txtareaDados.value ="";
            if(opcao==1){
                ifrIframe1.src= "http://sigeduca.seduc.mt.gov.br/ged/hwtmgedaluno.aspx?"+vetAluno[n]+",,HWMConAluno,DSP,1,0";
                ifrIframe1.addEventListener("load", coletaDados1);
            }
            if(opcao==2){
                ifrIframe1.src= "http://sigeduca.seduc.mt.gov.br/ged/hwtmgedaluno1.aspx?"+vetAluno[n]+",,HWMConAluno,DSP,0,1,0,1";
                ifrIframe1.addEventListener("load", coletaDados2);
            }
        }
    }
}

//Extrai dados da ABA "PESSOAL"
function coletaDados1() {
    if(n < vetAluno.length){
        //Dados gerais do Aluno
        a = a + vetAluno[n] +";"; cabecalho = "Cod Aluno;"; //Cod Aluno
        a = a + parent.frames[0].document.getElementById('span_CTLGEDALUIDINEP').innerHTML +";"; cabecalho = cabecalho+"MAT INEP;"; //Matrícula INEP
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNOM').innerHTML +";"; cabecalho = cabecalho+"Nome Aluno;";
        a = a + parent.frames[0].document.getElementById('CTLGERPESCPF').innerHTML.replace(/^([\d]{3})([\d]{3})([\d]{3})([\d]{2})$/, "$1.$2.$3-$4") +";"; cabecalho = cabecalho+"CPF do Aluno;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESRG').innerHTML +";"; cabecalho = cabecalho+"RG do aluno;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERORGEMICOD').innerHTML +";"; cabecalho = cabecalho+"Órgão Expedidor;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESSEXO').innerHTML +";"; cabecalho = cabecalho+"Sexo do Aluno;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESDTANASC').innerHTML +";"; cabecalho = cabecalho+"Data Nascimento;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNATDSC').innerHTML +";"; cabecalho = cabecalho+"Naturalidade;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNATUF').innerHTML +";"; cabecalho = cabecalho+"UF;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNOMMAE').innerHTML +";"; cabecalho = cabecalho+"Filiação 1;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNOMPAI').innerHTML +";"; cabecalho = cabecalho+"filiação 2;";

        //Contatos responsável 1
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNOMRESP').innerHTML+";"; cabecalho = cabecalho+"Nome do responsável 1;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESRESPCPF').innerHTML.replace(/^([\d]{3})([\d]{3})([\d]{3})([\d]{2})$/, "$1.$2.$3-$4")+";"; cabecalho = cabecalho+"CPF do responsável 1;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELRESDDDRESP').innerHTML+")"; //DDD Residencial
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELRESRESP').innerHTML+";"; cabecalho = cabecalho+"Tel Res Resp 1;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCELDDDRESP').innerHTML+")"; //DDD Celular
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCELRESP').innerHTML+";"; cabecalho = cabecalho+"Tel Celular Resp 1;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCOMDDDRESP').innerHTML+")"; //DDD Comercial
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCOMRESP').innerHTML+";"; cabecalho = cabecalho+"Tel Comercial Resp 1;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCONDDDRESP').innerHTML+")"; //DDD Contato
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCONRESP').innerHTML+";"; cabecalho = cabecalho+"Tel Contato Resp 1;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESEMAILRESP').innerHTML+";"; cabecalho = cabecalho+"E-mail Resp 1;";

        //Contatos responsável 2
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNOMRESP2').innerHTML+";"; cabecalho = cabecalho+"Nome do responsável 2;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESRESPCPF2').innerHTML.replace(/^([\d]{3})([\d]{3})([\d]{3})([\d]{2})$/, "$1.$2.$3-$4")+";"; cabecalho = cabecalho+"CPF do responsável 2;";
        a = a + "("+parent.frames[0].document.getElementById('CTLGERPESTELRESDDDRESP2').innerHTML+")"; //DDD Residencial
        a = a + parent.frames[0].document.getElementById('CTLGERPESTELRESRESP2').innerHTML+";"; cabecalho = cabecalho+"Tel Res Resp 2;";
        a = a + "("+parent.frames[0].document.getElementById('CTLGERPESTELCELDDDRESP2').innerHTML+")"; //DDD Celular
        a = a + parent.frames[0].document.getElementById('CTLGERPESTELCELRESP2').innerHTML+";"; cabecalho = cabecalho+"Tel Celular Resp 2;";
        a = a + "("+parent.frames[0].document.getElementById('CTLGERPESTELCOMDDDRESP2').innerHTML+")"; //DDD Comercial
        a = a + parent.frames[0].document.getElementById('CTLGERPESTELCOMRESP2').innerHTML+";"; cabecalho = cabecalho+"Tel Comercial Resp 2;";
        a = a + "("+parent.frames[0].document.getElementById('CTLGERPESTELCONDDDRESP2').innerHTML+")"; //DDD Contato
        a = a + parent.frames[0].document.getElementById('CTLGERPESTELCONRESP2').innerHTML+";"; cabecalho = cabecalho+"Tel Contato Resp 2;";
        a = a + parent.frames[0].document.getElementById('CTLGERPESEMAILRESP2').innerHTML+";"; cabecalho = cabecalho+"E-mail Resp 2;";

        //Contato da seção final da página (ENDEREÇO)
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELRESDDD').innerHTML+")"; //DDD Residencial 2
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELRES').innerHTML+";"; cabecalho = cabecalho+"Tel Residencial;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCELDDD').innerHTML+")"; //DDD Celular 2
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCEL').innerHTML+";"; cabecalho = cabecalho+"Tel Celular;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCOMDDD').innerHTML+")"; //DDD Comercial 2
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCOM').innerHTML+";"; cabecalho = cabecalho+"Tel Comercial;";
        a = a + "("+parent.frames[0].document.getElementById('span_CTLGERPESTELCONDDD').innerHTML+")"; //DDD Contato 2
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESTELCON').innerHTML+";"; cabecalho = cabecalho+"Tel Contato;";

        //Endereço
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESEND').innerHTML+";"; cabecalho = cabecalho+"Endereco Rua;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNMRLOG').innerHTML+";"; cabecalho = cabecalho+"Numero;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESCMPLOG').innerHTML+";"; cabecalho = cabecalho+"Complemento;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESBAIRRO').innerHTML+";"; cabecalho = cabecalho+"Bairro;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESENDCIDDSC').innerHTML+";"; cabecalho = cabecalho+"Cidade;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESENDUF').innerHTML+";"; cabecalho = cabecalho+"UF;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESCEP').innerHTML+";"; cabecalho = cabecalho+"CEP;";
        a = a + "\n";

        txtareaDados.value = cabecalho+"\n"+a;
        n=n+1;
        if(n < vetAluno.length){
            ifrIframe1.src = "http://sigeduca.seduc.mt.gov.br/ged/hwtmgedaluno.aspx?"+vetAluno[n]+",,HWMConAluno,DSP,1,0";
        }
    }
    if(n >= vetAluno.length){
        //navigator.clipboard.writeText(txtareaDados.value); ERRO
        txtareaDados.select();
        alert('finalizado');
    }
}

//Extrai dados da ABA "SOCIAL"
function coletaDados2() {
    if(n < vetAluno.length){
        a = a + vetAluno[n] +";"; cabecalho = "Cod Aluno;" //Cod Aluno
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNUMCARTAOSUS').innerHTML +";"; cabecalho = cabecalho+"Nro SUS;";
        a = a + parent.frames[0].document.getElementById('span_CTLGERPESNIS').innerHTML +";"; cabecalho = cabecalho+"Nro NIS;";
        a = a + parent.frames[0].document.getElementById('span_CTLGEDALUTIPOSANGUINEO').innerHTML.replace('SELECIONE', 'não consta')+";"; cabecalho = cabecalho+"Tipo sanguíneo;";
        a = a + parent.frames[0].document.getElementById('span_CTLGEDALURECATEEDUESP').innerHTML +";"; cabecalho = cabecalho+"Recebe Atendimento Especializado;";
        a = a + "\n";

        txtareaDados.value = cabecalho+"\n"+a;
        n=n+1;
        if(n < vetAluno.length){
            ifrIframe1.src = "http://sigeduca.seduc.mt.gov.br/ged/hwtmgedaluno1.aspx?"+vetAluno[n]+",,HWMConAluno,DSP,0,1,0,1";
        }
    }
    if(n >= vetAluno.length){
        navigator.clipboard.writeText(txtareaDados.value);
        txtareaDados.select();
        alert('finalizado');
    }
}

//BOTÃO EXIBIR ou ESCONDER
var exibir = '$("#credito1").slideToggle();if(this.value=="ESCONDER"){this.value="EXIBIR"}else{this.value="ESCONDER"}';
var btnExibir = document.createElement('input');
    btnExibir.setAttribute('type','button');
    btnExibir.setAttribute('id','exibir1');
    btnExibir.setAttribute('value','ESCONDER');
    btnExibir.setAttribute('class','menuSCT');
    btnExibir.setAttribute('style','background:#FF3300; width: 187px; border: 1px solid rgb(0, 0, 0); position: fixed; z-index: 2002; bottom: 0px; right: 30px;');
    btnExibir.setAttribute('onmouseover', 'this.style.backgroundColor = "#FF7A00"');
    btnExibir.setAttribute('onmouseout', 'this.style.backgroundColor = "#FF3300"');
    btnExibir.setAttribute('onmousedown', 'this.style.backgroundColor = "#EB8038"');
    btnExibir.setAttribute('onmouseup', 'this.style.backgroundColor = "#FF7A00"');
    btnExibir.setAttribute('onclick', exibir);
document.getElementsByTagName('body')[0].appendChild(btnExibir);

//DIV principal (corpo)
var divCredit = document.createElement('div');
    divCredit.setAttribute('id','credito1');
    divCredit.setAttribute('name','credito2');
    divCredit.setAttribute('class','menuSCT');
    divCredit.setAttribute('style','background: #DBDBDB; color: #000; width: 380px; text-align: center;font-weight: bold;position: fixed;z-index: 2002;padding: 5px 0px 0px 5px;bottom: 24px;right: 30px;height: 400px;');
document.getElementsByTagName('body')[0].appendChild(divCredit);

//Iframe
var ifrIframe1 = document.createElement("iframe");
ifrIframe1.setAttribute("id","iframe1");
ifrIframe1.setAttribute("src","about:blank");
ifrIframe1.setAttribute("style","height: 100px; width: 355px; display:none");
divCredit.appendChild(ifrIframe1);

//TEXTO CÓDIGO ALUNO
var textCodAluno = document.createTextNode("CÓDIGO DO ALUNO");
divCredit.appendChild(textCodAluno);

//textarea alunos a serem pesquisados
var txtareaAluno = document.createElement('TEXTAREA');
txtareaAluno.setAttribute('name','txtAluno');
txtareaAluno.setAttribute('id','txtAluno');
txtareaAluno.setAttribute('value','');
txtareaAluno.setAttribute('style','border:1px solid #000000;width: 355px;height: 82px; resize: none');
divCredit.appendChild(txtareaAluno);

//DIV NIS1
var divNIS = document.createElement('div');
divNIS.setAttribute('id','divNIS1');
divNIS.setAttribute('name','divNIS2');
divCredit.appendChild(divNIS);

//BOTÃO COLETAR DADOS PESSOAIS
var btnColetar1 = document.createElement('input');
btnColetar1.setAttribute('type','button');
btnColetar1.setAttribute('name','btnColetar1');
btnColetar1.setAttribute('value','Extrair de aba \"Pessoais\"');
btnColetar1.setAttribute('class','botaoSCT');
divCredit.appendChild(btnColetar1);
btnColetar1.onclick = function(){coletar(1)};

//BOTÃO COLETAR DADOS SOCIAIS
var btnColetar2 = document.createElement('input');
btnColetar2.setAttribute('type','button');
btnColetar2.setAttribute('name','btnColetar2');
btnColetar2.setAttribute('value','Extrair de aba \"Sociais\"');
btnColetar2.setAttribute('class','botaoSCT');
divCredit.appendChild(btnColetar2);
btnColetar2.onclick = function(){coletar(2)};

//QUEBRA LINHA
var quebraLinha1 = document.createElement("br");
divCredit.appendChild(quebraLinha1);
quebraLinha1 = document.createElement("br");
divCredit.appendChild(quebraLinha1);

//TEXTO DADOS COLETADOS
var textColetados = document.createTextNode("DADOS COLETADOS");
divCredit.appendChild(textColetados);

//textarea pra dados coletados
var txtareaDados = document.createElement('TEXTAREA');
txtareaDados.setAttribute('name','txtDados');
txtareaDados.setAttribute('id','txtDados');
txtareaDados.setAttribute('value','');
txtareaDados.setAttribute('style','border:1px solid #000000;width: 355px;height: 150px; resize: none');
divCredit.appendChild(txtareaDados);

//BOTAO SALVAR EM TXT
var btnSalvarTxt = document.createElement('input');
btnSalvarTxt.setAttribute('type','button');
btnSalvarTxt.setAttribute('name','btnSalvarTxt');
btnSalvarTxt.setAttribute('value','Salvar em CSV(Excel)');
btnSalvarTxt.setAttribute('class','botaoSCT');
divCredit.appendChild(btnSalvarTxt);
btnSalvarTxt.onclick = saveTextAsFile;

//DIV CREDITO
var divCredito = document.createElement('div');
divCredit.appendChild(divCredito);

var br1 = document.createElement('br');
divCredito.appendChild(br1);

var span1 = document.createElement('span');
span1.innerHTML = '>>Roberson Arruda<<';
divCredito.appendChild(span1);

br1 = document.createElement('br');
span1.appendChild(br1);

span1 = document.createElement('span');
span1.innerHTML = '(robersonarruda@outlook.com)';
divCredito.appendChild(span1);

br1 = document.createElement('br');
span1.appendChild(br1);

span1 = document.createElement('span');
span1.innerHTML = 'Extrator de Contatos';
divCredito.appendChild(span1);

window.scrollTo(0, document.body.scrollHeight);
