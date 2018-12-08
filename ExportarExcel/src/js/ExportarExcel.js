// Patrick de Souza -- 04/12/2018
var ExportaExcel = {
    //Metodo para gerar o excel 
    //data deve ser enviada em JSON
    //filename nome do arquivo
    Exportar : function(data, nomeArquivo){
        //Exibir coluna como true mas caso necessário enviar como paramentro
        ExportaExcel.JSONParaExcel(data, nomeArquivo,true);
    }
    , JSONParaExcel: function(JSONData, nomeArquivo, exibirColuna ){

        var data        = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;
        if(data[0] === undefined){
        	var aux = [];
        	aux.push(data);
        	data = aux;
        }

        //Function para download do excel encapsulada
        var saveData    = (function () {

            // variavel "a" equivalente a tag <a> do HTML
            var a = document.createElement("a");
            document.body.appendChild(a);
            a.style = "display: none";

            return function (data, nomeArquivo) {
               
                var blob    = new Blob([data], { type: "application/vnd.ms-excel" }),
                url         = window.URL.createObjectURL(blob);
                a.href      = url;
                a.download  = nomeArquivo;
                a.click();
                window.URL.revokeObjectURL(url);
            };
        }());

        nomeArquivo    = nomeArquivo + ".csv";
        //Metodo para gerar o CSV para JSON
        data        = ExportaExcel.JSONParaCSV(data, nomeArquivo, exibirColuna);
        //Metodo para fazer o download do excel
        if(data == "")
            return;
        //Realizar download do excel
        saveData(data, nomeArquivo);
    }
    , JSONParaCSV : function (JSONData, Titulo, exibirColuna) {
        
        //Json 
        var data    = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;
        //Conteudo
        var CSV     = '';    
        //CSV         += Titulo + '\r\n\n';

        //Exibir coluna no excel 
        if (exibirColuna) {
            var linha = "";
       
            //Aqui eu pego a primeira coluna do excel para extrair a coluna
            for (var index in data[0]) {
           
                //Aqui eu faço concatenação com ponto e virgula
                linha += index + ';';
            }

            linha = linha.slice(0, -1);
       
            //adiciono no CSV
            CSV += linha + '\r\n';
        }
   
        //1º Loop para pecorrer cada registro
        for (var i = 0; i < data.length; i++) {
            var linha = "";
       
            //2º Loop para extrair cada coluna 
            for (var index in data[i]) {
                //TODO: Validar o uso de expressao regular 
                if (typeof data[i][index] === 'string') {
                    var mapaAcentosHex 	= {
                        a : /[\xE0-\xE6]/g,
                        A : /[\xC0-\xC6]/g,
                        e : /[\xE8-\xEB]/g,
                        E : /[\xC8-\xCB]/g,
                        i : /[\xEC-\xEF]/g,
                        I : /[\xCC-\xCF]/g,
                        o : /[\xF2-\xF6]/g,
                        O : /[\xD2-\xD6]/g,
                        u : /[\xF9-\xFC]/g,
                        U : /[\xD9-\xDC]/g,
                        c : /\xE7/g,
                        C : /\xC7/g,
                        n : /\xF1/g,
                        N : /\xD1/g
                    };
                    for(var letra in mapaAcentosHex ) {
                        var expressaoRegular = mapaAcentosHex[letra];
                        data[i][index] = data[i][index].replace(expressaoRegular, letra);
                    }  
                }
                linha += ('"' + data[i][index] + '";').replace("null","");
            }

            linha.slice(0, linha.length - 1);
       
            //adiciono no CSV
            CSV += linha + '\r\n';
        }

        if (CSV == '') {        
            alert("Falha na geração, dados inválidos");
            return CSV;
        }   
        return CSV;
    }
    , Carregar: function(){
        //Evento do click do exportar excel
        document.getElementById("exportar").addEventListener("click", ExportaExcel.ExportarBotao);
        //Evento do download do JS
        document.getElementById("download").addEventListener("click", ExportaExcel.DownloadJS);
    }
    , ExportarBotao: function(){
        var data = document.getElementById("JSON").value;
        var nomeArquivo = "Excel";
        if(data == "")
            return;
        ExportaExcel.Exportar(data,nomeArquivo);
    }
    , DownloadJS: function(){
        window.open("src/js/ExportarExcelDownload.js","download");
    }
};
