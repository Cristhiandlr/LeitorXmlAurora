document.getElementById('generateTableBtn').addEventListener('click', function () {
    const fileInput = document.getElementById('fileInput');
    const files = fileInput.files;
    if (files.length === 0) {
        alert('Selecione pelo menos um arquivo XML.');
        return;
    }

    const rows = []; // Para armazenar todas as linhas de dados

    // Adiciona o cabeçalho ao array
    const headers = [
        'CFOP', 'Número CTe', 'Data/Hora Emissão', 'Município Início', 'Município Fim',
        'Observações', 'Nome Empresa', 'Nome Cliente', 'CNPJ Cliente Final',
        'Nome Cliente Final', 'Chave Nota Fiscal', 'Peso Carga', 'Valor Carga', 
        'Valor do CTE', 'Valor Frete', 'Porcentagem do Valor'
    ];
    rows.push(headers);

    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = function (event) {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(event.target.result, 'text/xml');

            // Extrai as informações do XML
            const cfop = xmlDoc.querySelector('ide CFOP')?.textContent || '';
            const nCT = xmlDoc.querySelector('ide nCT')?.textContent || '';
            const dhEmi = xmlDoc.querySelector('ide dhEmi')?.textContent || '';
            const xMunIni = xmlDoc.querySelector('ide xMunIni')?.textContent || '';
            const xMunFim = xmlDoc.querySelector('ide xMunFim')?.textContent || '';
            const xObs = xmlDoc.querySelector('compl xObs')?.textContent || '';
            const nomeEmpresa = xmlDoc.querySelector('emit xNome')?.textContent || '';
            const nomeCliente = xmlDoc.querySelector('rem xNome')?.textContent || '';
            const cnpjClienteFinal = xmlDoc.querySelector('dest CNPJ')?.textContent || '';
            const nomeClienteFinal = xmlDoc.querySelector('dest xNome')?.textContent || '';
            const chaveNotaFiscal = xmlDoc.querySelector('infDoc chave')?.textContent || '';
            
            // Substitui pontos por vírgulas nas colunas numéricas
            let pesoCarga = (xmlDoc.querySelector('infCarga qCarga')?.textContent || '').replace('.', ',');
            let valorCarga = (xmlDoc.querySelector('infCarga vCarga')?.textContent || '').replace('.', ',');
            let valorCTE = (xmlDoc.querySelector('Comp vComp')?.textContent || '').replace('.', ',');

            // Converte os valores para números (substituindo vírgula por ponto) para cálculos
            pesoCarga = parseFloat(pesoCarga.replace(',', '.'));
            valorCTE = parseFloat(valorCTE.replace(',', '.'));

            // Calcula o Valor Frete (Peso Carga * 0.92) e converte para string com vírgula
            const valorFrete = (pesoCarga * 0.92).toFixed(2).replace('.', ',');

            // Calcula a Porcentagem do Valor (Valor CTE / Valor Frete)
            const porcentagemValor = (valorCTE / (pesoCarga * 0.92)).toFixed(2).replace('.', ',');

            // Adiciona os dados como uma linha ao array
            const row = [
                cfop, nCT, dhEmi, xMunIni, xMunFim, xObs, nomeEmpresa, nomeCliente,
                cnpjClienteFinal, nomeClienteFinal, chaveNotaFiscal, pesoCarga.toFixed(2).replace('.', ','),
                valorCarga, valorCTE.toFixed(2).replace('.', ','), valorFrete, porcentagemValor
            ];
            rows.push(row);

            // Verifica se todos os arquivos foram processados
            if (rows.length === files.length + 1) {
                // Cria um novo workbook e adiciona a planilha com os dados
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(rows);
                XLSX.utils.book_append_sheet(wb, ws, 'Dados XML');

                // Gera e baixa o arquivo Excel
                XLSX.writeFile(wb, 'Xml_Aurora.xlsx');
            }
        };
        reader.readAsText(file);
    });
});
