import JSZip from 'jszip';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

function numeroPorExtenso(numero) {
    const unidades = ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"];
    const especiais = ["dez", "onze", "doze", "treze", "catorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
    const dezenas = ["", "dez", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];

    if (numero < 10) return unidades[numero];
    if (numero < 20) return especiais[numero - 10];
    if (numero < 100) {
        let unidade = numero % 10;
        let dezena = Math.floor(numero / 10);
        return dezenas[dezena] + (unidade ? " e " + unidades[unidade] : "");
    }
    return "Número fora do intervalo suportado";
}

function gerarContrato() {
    console.log('gerarContrato called');
    
    let nome = document.getElementById('nome').value;
    let nacionalidade = document.getElementById('nacionalidade').value;
    let estado_civil = document.getElementById('estado_civil').value;
    let profissao = document.getElementById('profissao').value;
    let cpf = document.getElementById('cpf').value;
    let endereco = document.getElementById('endereco').value;
    let cep = document.getElementById('cep').value;
    let cidade = document.getElementById('cidade').value;
    let foro = document.getElementById('foro').value;
    let empresa = document.getElementById('empresa').value;

    // Obter a data atual
    let dataAtual = new Date();
    let dia = dataAtual.getDate();  // Change from 'data' to 'dia'
    let mes = dataAtual.toLocaleString('default', { month: 'long' });
    let ano = dataAtual.getFullYear();

    console.log('Fetching contract template...');
    fetch('contrato modelo.docx') // Update this line to fetch the new contract template
        .then(response => {
            if (!response.ok) {
                throw new Error('Erro ao carregar o modelo de contrato.');
            }
            console.log('Template fetched successfully');
            return response.arrayBuffer();
        })
        .then(arrayBuffer => {
            let zip;
            try {
                zip = new PizZip(arrayBuffer);
                console.log('PizZip created successfully');
            } catch (error) {
                console.error('Erro ao criar PizZip:', error);
                document.getElementById('status').innerText = "Erro ao processar o contrato.";
                return;
            }

            let doc;
            try {
                doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
                console.log('Docxtemplater created successfully');
            } catch (error) {
                console.error('Erro ao criar docxtemplater:', error);
                document.getElementById('status').innerText = "Erro ao processar o contrato.";
                return;
            }

            // Ensure data types are correct
            const templateData = {
                nome_autor: String(nome),
                nacionalidade_autor: String(nacionalidade),
                estadocivil_autor: String(estado_civil),
                profissao_autor: String(profissao),
                cpf_autor: String(cpf),
                endereco_autor: String(endereco),
                cep_autor: String(cep),
                cidade_autor: String(cidade),
                foro_autor: String(foro),
                empresa_autor: String(empresa),
                data: String(dia),  // Change from 'data' to 'dia'
                mes: String(mes),
                ano: String(ano),
            };

            console.log('Data set for template:', templateData);

            doc.setData(templateData);

            try {
                doc.render();
                console.log('Contract rendered successfully');
                let out = doc.getZip().generate({
                    type: "blob",
                    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                });
                saveAs(out, `contrato_de_${nome}.docx`);
                document.getElementById('status').innerText = "Contrato gerado com sucesso!";
            } catch (error) {
                console.error('Erro ao renderizar o contrato:', error);
                document.getElementById('status').innerText = "Erro ao gerar o contrato.";
            }
        })
        .catch(error => {
            console.error('Erro no processamento do fetch:', error);
            document.getElementById('status').innerText = "Erro ao carregar o modelo de contrato.";
        });
}

// Expose the function to the global scope
window.gerarContrato = gerarContrato;
