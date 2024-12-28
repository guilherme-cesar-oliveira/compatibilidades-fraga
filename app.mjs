import fetch from 'node-fetch';
import ExcelJS from 'exceljs'


let logins = [{ email: "Seu email de acesso", password: "Sua senha de acesso." }]


  // Exemplo de uso
  let ultimoIndiceSelecionado = -1; // Inicializado como -1 para garantir que a primeira seleção seja sempre diferente

  const planilhaPath = 'modelos_de_pecas.xlsx';
  const nomeDaAba = 'Planilha1';
  const nomeDaColuna = 'A';

  //Nome das planilhas que serão exportadas
  const outputFilePath = `Produtos Calpen - Infos Gerais.xlsx`;
  const outputFilePath2 = `Produtos Calpen - Referencias Cruzada.xlsx`;
  const outputFilePath3 = `Produtos Calpen - Compatibilidades.xlsx`;

  let start = false

  let token = await generatetoken(false)
  let accesstoken = token.access_token

  setInterval(async function() {
    if(start == true){
    token = await generatetoken();
    accesstoken = token.access_token
    console.log('Token Atualizado Com Sucesso')

    } 
}, 1200000); // Verifica a cada 20 minutos 1200000


// Função auxiliar para buscar o valor de uma especificação
function getSpecificationValue(specifications, description) {
  if (specifications && specifications.length > 0) {
      const spec = specifications.find(value => value.description === description);
      return spec ? spec.value : null;
  }
  return null;
}

function selecionarObjetoAleatorio() {
  // Exemplo de uso:
const arrayDeObjetos = logins
let novoIndiceAleatorio;
  
do {
  novoIndiceAleatorio = Math.floor(Math.random() * arrayDeObjetos.length);
} while (novoIndiceAleatorio === ultimoIndiceSelecionado); // Repete até encontrar um índice diferente

ultimoIndiceSelecionado = novoIndiceAleatorio; // Atualiza o último índice selecionado

return arrayDeObjetos[novoIndiceAleatorio];
}

async function lerColuna(planilhaPath, nomeDaAba, nomeDaColuna) {
    const workbook = new ExcelJS.Workbook();
  
    // Carrega a planilha
    await workbook.xlsx.readFile(planilhaPath);
  
    // Obtém a referência para a aba desejada
    const aba = workbook.getWorksheet(nomeDaAba);
  
    // Encontra a coluna desejada
    const colunaDesejada = aba.getColumn(nomeDaColuna);
  
    // Obtém os valores da coluna
    const valoresColuna = colunaDesejada.values.slice(2)

  
    return valoresColuna;
}
  
async function exportToXLSM(dataArray, outputFilePath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dados');
  
    // Add column headers
    console.log(dataArray)
    const headers = Object.keys(dataArray[0]);
    worksheet.addRow(headers);
  
    // Add data from the array to the Excel file
    dataArray.forEach((item) => {
      const row = [];
      headers.forEach((header) => {
        row.push(item[header]);
      });
      worksheet.addRow(row);
    });
  
    // Save the Excel file
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`XLSM file saved at: ${outputFilePath}`);
}

async function generatetoken(){

let selectedCredentials = selecionarObjetoAleatorio();

const myHeaders = new Headers();
myHeaders.append("Content-Type", "application/x-www-form-urlencoded");

const urlencoded = new URLSearchParams();
urlencoded.append("client_id", "catalog-api");
urlencoded.append("client_secret", "ea1a09ecd75cee3d5a5203ba0f2dd003");


urlencoded.append("grant_type", "password");
urlencoded.append("password", selectedCredentials.password);
urlencoded.append("username", selectedCredentials.email);



const requestOptions = {
  method: "POST",
  headers: myHeaders,
  body: urlencoded,
  redirect: "follow"
};

return fetch("https://accounts.fraga.com.br/realms/cat_pecamentor/protocol/openid-connect/token", requestOptions)
  .then((response) => response.json())
  .then((result) => {
    start = true
    return result
  })
  .catch((error) => console.error(error));

}

async function fetchData(pesquisa,token) {
    const endpoint = 'https://apiv2.catalogofraga.com.br/graphql/'; // Substitua pelo seu endpoint GraphQL    
    const query = `query GetProducts{
      catalogSearch(
          query: "${pesquisa}"
          take: 10
          skip: 0
          market: BRA
        ) {
          nodes {
            product {
              id
              partNumber
              brand {
                name
              }
              specifications {
                category
                important
                description
                value
              }
              images {
                category
                thumbnailUrl
              }
              applicationDescription
             
              status
              distributors {
                code
                distributor {
                  name
                }
              }
              containsUniversalApplication
            }
           
          }
          
        }
      }`;

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({ query })

        });

        let data = await response.json();

        if(data.data && data.data.catalogSearch && data.data.catalogSearch.nodes.length > 0 ){
        data = data.data.catalogSearch.nodes[0].product
        return data
        }else if(!data.data){
          console.log(data.errors)
        }
        
        
    } catch (error) {
        console.error('Ocorreu um erro:', error);
    }
}

async function fetchData2(id,token) {
    const endpoint = 'https://apiv2.catalogofraga.com.br/graphql/'; // Substitua pelo seu endpoint GraphQL    
    const query = `query GetProductById {
        product(id: "${id}", market: BRA) {
          id
          partNumber
          brand {
            name
            
          }
          applicationDescription
          images {
            imageUrl
            thumbnailUrl
            category
            
          }
          specifications {
            id
            category
            description
            value
            important
            
          }
          crossReferences {
            brand {
              name
              
            }
            partNumber
            
          }
          videos
          vehicles {
            brand
            name
            model
            engineName
            engineConfiguration
            endYear
            note
            only
            restriction
            startYear
            
          }
          components {
            partNumber
            productGroup
            applicationDescription
            activeCatalog
            status
            
          }
          distributors {
            code
            distributor {
              name
              
            }
            
          }
          status
          containsUniversalApplication
          
        }
      }`;

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({ query })

        });

        let data = await response.json();
        return data.data.product
        
        
    } catch (error) {
        console.error('Ocorreu um erro:', error);
    }
}

let ao = []
let ao2 = []
let ao3 = []

function removerThumb(url) {
  return url.replace(/thumb-/i, ''); // A expressão regular /thumb-/i vai corresponder ao termo "thumb-" independentemente de ser maiúsculas ou minúsculas e substituí-lo por uma string vazia.
}

function formatartitulo(titulo) {
  const termos = titulo.split(' '); // Divide o título em termos separados por espaço
  if (termos.length >= 2) { // Verifica se existem pelo menos dois termos
      termos.pop(); // Remove o último termo
      termos.pop(); // Remove a virgula
      termos.pop(); // Remove o penúltimo termo

      return termos.join(' '); // Junta os termos restantes em uma única string
  } else {
      return titulo; // Se houver menos de dois termos, retorna o título original
  }
}

function calcularSimilaridade(str1, str2) {
  str1 = String(str1);
  str2 = String(str2);

  const set1 = new Set(str1.split(''));
  const set2 = new Set(str2.split(''));

  const intersection = new Set([...set1].filter(x => set2.has(x)));
  const union = new Set([...set1, ...set2]);

  const similaridade = intersection.size / union.size;
  return similaridade;
}

async function main() {


   let ids = await lerColuna(planilhaPath, nomeDaAba, nomeDaColuna)
    let count = 0
    for (const id of ids) {
        // Encapsule o setTimeout em uma Promise para aguardar o atraso
       count+=1
        console.log(count)
        let a = await fetchData(id,accesstoken);

      if(a){

        let productid = a.id
        let co = await fetchData2(productid,accesstoken)

        let cop = ao.find(value=>{
          if(value.Id == productid){
              console.log('Já esta no array')
          return value.Id}})

          if(cop == undefined){

          let refs 
          let vehicles = null
          let brand = a.brand.name
          let partNumber = a.partNumber
          let brand2 = null
          let partnumber2 = null
          let substituido = false;

          if(co && co.vehicles != null){
          vehicles = co.vehicles
          }


        if(co && co.crossReferences != null){

          refs = co.crossReferences
      
        refs.forEach(ref =>{
          let refobj = {}

          brand2 = ref.brand.name
          partnumber2 = ref.partNumber
  
          let similaridade = calcularSimilaridade(id, partnumber2)

          if(similaridade >= 0.80 && substituido == false){
          //Se a similaridade for alta, substitui a marca e partnumber do arquivo Infos Gerais por aquelas que são similares aos modelos informados.

            brand = ref.brand.name
            partNumber = ref.partNumber
            substituido = true; // Define a flag para indicar que as variáveis foram atualizadas

          
          }

          if(similaridade >= 0.80){
  
              brand2 = a.brand.name
              partnumber2 = a.partNumber
  
            }

            refobj.brand = brand2
            refobj.partNumber = partnumber2
            refobj.id = productid

            ao2.push(refobj)

        })

        }
        if(vehicles != null){
        vehicles.forEach(vei =>{
          let veiobj = {}
            veiobj.id = productid
            veiobj.brand =vei.brand
            veiobj.name =vei.name
            veiobj.model =vei.model
            veiobj.engineName =vei.engineName
            veiobj.engineConfiguration =vei.engineConfiguration
            veiobj.endYear =vei.endYear
            veiobj.note =vei.note
            veiobj.only =vei.only
            veiobj.restriction =vei.restriction
            veiobj.startYear =vei.startYear

            ao3.push(veiobj)
        })
        }

      let thumb = (a.images)? a.images[0].thumbnailUrl:null
      let titulo = formatartitulo(a.applicationDescription)
      if(thumb!=null){thumb = removerThumb(thumb)}

        const o = {
          Id: a.id,
          Id: productid,
          Descrição: titulo,
          Part_Number: partNumber,
          Marca: brand,
          Ncm: getSpecificationValue(a.specifications, 'NCM'),
          Origem: getSpecificationValue(a.specifications, 'Origem do produto'),
          Peso: getSpecificationValue(a.specifications, 'Peso bruto'),
          Peso_Liquido: getSpecificationValue(a.specifications, 'Peso líquido'),
          Altura: getSpecificationValue(a.specifications, 'Altura'),
          Largura: getSpecificationValue(a.specifications, 'Largura'),
          Comprimento: getSpecificationValue(a.specifications, 'Comprimento'),
          Diametro_Interno: getSpecificationValue(a.specifications, 'Diâmetro interno'),
          Diametro_Externo: getSpecificationValue(a.specifications, 'Diâmetro externo'),
          Altura_Embalagem: getSpecificationValue(a.specifications, 'Altura da embalagem'),
          Largura_Embalagem: getSpecificationValue(a.specifications, 'Largura da embalagem'),
          Comprimento_Embalagem: getSpecificationValue(a.specifications, 'Comprimento da embalagem'),
          Thumb: thumb,
      };


        ao.push(o);
    }
  }   
}
}

await main()
//Caso for fazer algum processo muito grande, comentar linhas, o programa pode crashar se o tamanho dos arrays for muito grande, uma solução possivel é usar o db para não perder todos os dados caso crashe
exportToXLSM(ao, outputFilePath);
if(ao2.length > 0){exportToXLSM(ao2, outputFilePath2);}
exportToXLSM(ao3, outputFilePath3);


