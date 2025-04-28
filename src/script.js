let pyodide;
let filesContent = {};

// Inicializa o Pyodide
async function loadPyodideAndPackages() {
    pyodide = await loadPyodide();
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("pandas"); // Para gerar DataFrame
    await micropip.install("openpyxl"); // Para gerar Excel
    console.log("Pyodide e pacotes carregados.");
    document.getElementById('log').textContent = "Pyodide carregado.\n";
}

// Carrega arquivos XML quando a pasta é selecionada
document.getElementById('folderInput').addEventListener('change', async function(event) {
    const files = event.target.files;
    filesContent = {};
    document.getElementById('status').textContent = `Lendo ${files.length} arquivos...`;
    document.getElementById('log').textContent += `Lendo ${files.length} arquivos...\n`;
    
    for (let file of files) {
        const text = await file.text();
        filesContent[file.webkitRelativePath] = text;
    }

    document.getElementById('generateButton').disabled = false;
    document.getElementById('status').textContent = `${files.length} arquivos carregados. Pronto para gerar.`;
    document.getElementById('log').textContent += `${files.length} arquivos carregados.\n`;
});

// Função para gerar a planilha
async function generateSpreadsheet() {
    document.getElementById('status').textContent = "Processando...";
    document.getElementById('generateButton').disabled = true;
    document.getElementById('log').textContent += "Iniciando processamento...\n";

    try {
        // Converte filesContent em um array de pares [file_path, content]
        const filesContentArray = Object.entries(filesContent);
        pyodide.globals.set("files_content", filesContentArray);

        // Configura o redirecionamento de stdout para o log
        pyodide.runPython(`
from js import document
import sys
class LogWriter:
    def write(self, text):
        document.getElementById('log').textContent += text
    def flush(self):
        pass
sys.stdout = LogWriter()
        `);

        // Código Python para processar XMLs e gerar Excel
        const pythonCode = `
import xml.etree.ElementTree as ET
import pandas as pd
import io
import openpyxl

def process_xmls(files_content):
    # Tags específicas a serem extraídas
    tags_especificas = ["CPF", "CNPJ", "xNome", "xLgr", "nro", "CEP", "IE"]
    
    # Namespace utilizado no XML
    namespace = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    
    # Lista para armazenar os dados
    dados = []
    
    print(f"Total de arquivos recebidos: {len(files_content)}")
    for file_path, content in files_content:
        print(f"Processando: {file_path}")
        try:
            # Parseia o XML a partir do conteúdo
            root = ET.fromstring(content)
            
            # Procurar a tag <dest> dentro do namespace
            dest = root.find(".//nfe:dest", namespace)
            if dest is not None:
                linha = {}
                # Acessar as informações de endereços dentro de <enderDest>
                enderDest = dest.find("nfe:enderDest", namespace)
                
                # Buscar as tags dentro de <dest> e <enderDest>
                for tag in tags_especificas:
                    # Definir o elemento de acordo com a tag
                    if tag in ["xLgr", "nro", "CEP"]:
                        elemento = enderDest.find(f"nfe:{tag}", namespace) if enderDest is not None else None
                    else:
                        elemento = dest.find(f"nfe:{tag}", namespace)
                    
                    # Verifica se a tag foi encontrada
                    linha[tag] = elemento.text if elemento is not None else None
                dados.append(linha)
                print(f"Dados extraídos de {file_path}: {linha}")
            else:
                print(f"Tag <dest> não encontrada em {file_path}")
        except Exception as e:
            print(f"Erro ao processar {file_path}: {str(e)}")
    
    print(f"Total de linhas geradas: {len(dados)}")
    # Cria um DataFrame
    df = pd.DataFrame(dados)
    
    # Salva o DataFrame como Excel em um buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Executa o processamento
result = process_xmls(files_content)
result
        `;

        // Executa o código Python
        const excelContent = await pyodide.runPythonAsync(pythonCode);

        // Cria o download do arquivo Excel
        const blob = new Blob([new Uint8Array(excelContent)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'informacoesDestinatario.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        document.getElementById('status').textContent = "Planilha gerada e baixada com sucesso!";
        document.getElementById('log').textContent += "Planilha gerada e baixada.\n";
    } catch (error) {
        console.error(error);
        document.getElementById('status').textContent = "Erro ao processar os arquivos: " + error.message;
        document.getElementById('log').textContent += `Erro: ${error.message}\n`;
    } finally {
        document.getElementById('generateButton').disabled = false;
    }
}

// Carrega o Pyodide ao iniciar a página
loadPyodideAndPackages();