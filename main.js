var tabela = [], aux = [], aux2 = []

function lerArquivo(arquivo, tipo = 0){
    if ((arquivo[0].name).includes('.xlsx')) {

        var leitor = new FileReader();
        // leitor.readAsArrayBuffer(arquivo);
        leitor.onload = function(e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
              type: 'binary'
            });
            
            if(tipo == 0){
                aux = pegarPlanilhaManse(workbook)
            }else{
                aux = pegarPlanilhaSohee(workbook)
            }

            aux.sort(function compare(a, b) {
                if (a.id < b.id) return -1;
                if (a.id > b.id) return 1;
                return 0;
            })

            removeRegistrosDuplos()

            somaDadosExistentes()

            insereDadosNovos()

            tabela.sort(function compare(a, b) {
                if (a.id < b.id) return -1;
                if (a.id > b.id) return 1;
                return 0;
            })

            Swal.fire({
                title: 'Sucesso',
                text: 'Arquivo de excel foi incluido para a mescla',
                type: 'success',
                confirmButtonText: 'Confirmar'
            })

        }

        leitor.readAsBinaryString(arquivo[0]);
    } else {
        Swal.fire({
            title: 'Error!',
            text: 'Favor enviar arquivos excel no formato .xlxs',
            type: 'error',
            confirmButtonText: 'Confirmar'
        })
    }          
}

function ExportarPlanilha () {
    var hoje = new Date(2017,12,19);
    var wb =  XLSX.utils.book_new();
    wb.Props = {
        Title:"titulo",
        Subject:"assunto",
        Author:"autor",
        CreatedDate: hoje
    };
    wb.SheetNames.push("test sheet");
    var ws = XLSX.utils.json_to_sheet(tabela);
    wb.Sheets["test sheet"] = ws
    let exportExcel = XLSX.write(wb, {bookType:'xlsx', type:'binary'});

    function s2ab(s){
        let buf = new ArrayBuffer(s.length);
        let view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf
    }

    saveAs(new Blob([s2ab(exportExcel)], {type: "application/octet-stream"}), "PlanilhaMesclada_"+hoje.getDate()+"_"+hoje.getMonth()+1+"_"+hoje.getFullYear()+".xlsx");
}

function incluiValorCampo(folha, modelo, coluna, pos, campo){
    if(folha[coluna+pos] != undefined){
        modelo[campo] = folha[coluna+pos].w;
    }else{
        modelo[campo] = 0;
    };
}

function pegarPlanilhaManse(wb){

    let folha = wb.Sheets[wb.SheetNames[0]],
        pos = 2,
        modelo = {},
        tabela = []

    while(folha['A'+pos] != undefined){
        modelo = {}
        modelo.id  = folha['A'+pos].w.replace('-','').slice(0,5)

        incluiValorCampo(folha, modelo, 'B', pos, "Descricao")
        incluiValorCampo(folha, modelo, 'F', pos, "Total")

        tabela = [...tabela, modelo]
        pos++
    }
    return tabela;
}

function pegarPlanilhaSohee(wb){

    let folha = wb.Sheets[wb.SheetNames[0]],
        pos = 2,
        modelo = {},
        tabela = []

    while(folha['A'+pos] != undefined){
        modelo = {}
        modelo.id  = folha['A'+pos].w.replace('-','').slice(0,5)

        incluiValorCampo(folha, modelo, 'B', pos, "Descricao")
        incluiValorCampo(folha, modelo, 'H', pos, "Total")

        tabela = [...tabela, modelo]
        pos++
    }
    return tabela;
}

function formatarSoma(a, b){
    return parseInt(a, 10) + parseInt(b, 10)
}

function removeRegistrosDuplos(){
    var novoRegistro=[], novoElemento = {}
    for (let i = 0; i < aux.length; i++) {
        if(novoElemento==''){
            novoElemento = aux[i]
        }
        if(i+2 <= aux.length)
        {
            if(aux[i].id == aux[i+1].id){
                novoElemento.Total == undefined ? novoElemento.Total = aux[i+1].Total : novoElemento.Total = formatarSoma(novoElemento.Total, aux[i+1].Total)
            }else{
                if(novoElemento.id!=undefined){novoRegistro.push(novoElemento)}
                novoElemento = aux[i+1]
            }
        }
    }
    aux = novoRegistro
}

function somaDadosExistentes(){
    for (let j = 0; j < tabela.length; j++) {
        for (let i = 0; i < aux.length; i++) {
            if(aux[i].id == tabela[j].id){
                tabela[j].Total = formatarSoma(tabela[j].Total, aux[i].Total)
            }
        }
    }
}

function insereDadosNovos() {
    for (let i = 0; i < aux.length; i++) {
        var encontrou = false
        for (let j = 0; j < tabela.length; j++) {
            if(aux[i].id == tabela[j].id){
                encontrou = true
            }
        }
        if(encontrou == false){
            tabela = [...tabela, aux[i]]
        }
    }
}