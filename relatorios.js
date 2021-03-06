function writeReport(data) {  
  const initRowNumber = 1;
  const initColumnNumber = 1;
  const numberOfRows = data.length
  const numberOfColumns = data[0].length
  console.log("writeReport => numberOfColumns", numberOfColumns)

  
 
  

}


function getFullReportData() {

  const cardsByDate = getCardsbyDate(getDatesFromReport())
  const cardsPreparado = getDataFromPreparadoSheet()
  const cardsDetails = getDetailedCardData()
  const cardsProjetos = getDataFromProjetosSheet()

  let totalArray = []

  cardsByDate.map((card, index) => {

    totalArray = [...totalArray, 
      {
        ...card,
        ...getPreparadoInfo(card.id, cardsPreparado),
        ...getCardDetails(card.id, cardsDetails),
        ...getProjectDetails(card.projeto, cardsProjetos),
      }]
  })
  console.log(`total ${JSON.stringify(totalArray[0])}`)
  return totalArray
}


function getCardsbyDate(dates) {

  if (!dates) return null;

  const cards = getDataFromDoneSheet()

  const cardsFiltradosPorData = cards.filter(card => card[9] > dates[0] && card[9] < dates[1]).map((filteredCard) => {
    return {
      "id": filteredCard[0],
      "atestadoPor": filteredCard[2],
      "atestadoEm": filteredCard[9],
      "textoAtestado": filteredCard[6],
      "projeto":filteredCard[1]
    }
  })
  //console.log(`#FiltradosPorData por data: ${JSON.stringify(cardsFiltradosPorData)}`)

  return cardsFiltradosPorData
}

function getPreparadoInfo(cardId, dataFromPreparadoSheet) {
  if (!cardId || !dataFromPreparadoSheet) return
  const preparadoData = dataFromPreparadoSheet.find(card => card[0] === cardId);
  //console.log(preparadoData)
  return preparadoData ? {
    "id": preparadoData[0],
    "autorizadoPor": preparadoData[2],
    "autorizadoEm": preparadoData[9],
    "textoAutorizado": preparadoData[7]
  } : undefined;
}


function getCardDetails(cardId, detailedCardData) {
  if (!cardId || !detailedCardData) return
  const detailedData = detailedCardData.find(card => card[2] === cardId);
  //console.log(detailedData);
  return detailedData ? {
    "id": detailedData[2],
    "timeSpentSeconds": detailedData[8] / 3600,
    "horasFaturar": detailedData[9],
    "summary": detailedData[3]
  } : undefined;
}

function getProjectDetails(projectName, projectDetails) {
  if (!projectName || !projectDetails) return
  const detailedData = projectDetails.find(card => card[8] === projectName);
  //console.log(detailedData); demandante	entidade	projeto	time	perc_cni	perc_sesi	perc_senai	perc_iel	fornecedor
  return detailedData ? {
    "demandante": detailedData[6],
    "entidade": detailedData[7],
    "projeto": detailedData[8],
    "time": detailedData[9],
    "cni": detailedData[10],
    "sesi": detailedData[11],
    "senai": detailedData[12],
    "iel": detailedData[13],
    "fornecedor": detailedData[14],

  } : undefined;
}


function testes() {
  // console.log(`teste 1: ${getCardDetails(null,getDetailedCardData())}`)
  // console.log(`teste 2: ${getCardDetails("",getDetailedCardData())}`)
  // console.log(`teste 3: ${getCardDetails("UN-13",getDetailedCardData())}`)
  // console.log(`teste 4: ${getCardDetails("UN-13",null)}`)
  console.log(`teste 5: ${JSON.stringify(getProjectDetails("Unind??stria", getDataFromProjetosSheet()))}`)

  //getCardsbyDate(getDatesFromReport())
}

function getDadosCompletos() {

  const dados_done = getCardsbyDate(getDatesFromReport())
  const dados_prep = getDataFromPreparadoSheet()
  const dados_card = getDetailedCardData()
  const dados_projetos = getDataFromProjetosSheet()

  let dados_final = []

  dados_done.forEach(function (item, index, arr) {// para cada item filtrado por data, complementar a informa????o de preparado

    const dados_complementares = dados_prep.find(function (card) {

      return card[0] === item[0]

    })
    item.push(dados_complementares[2], dados_complementares[11], dados_complementares[10])


    const horasDesc = dados_card.find(function (card) {


      return card[2] === item[0]
    })
    console.log(`HorasDesc: ${horasDesc}`)
    item.push(horasDesc[8] / 60 / 60, horasDesc[9], horasDesc[3])
    //console.log(`HorasDesc: ${horasDesc[8] / 60 / 60} / ${horasDesc[9]}`)



    const projetos = dados_projetos.find(function (projeto) {

      return projeto[8] === item[10]

    })
    item.push(projetos[6], projetos[7], projetos[8], projetos[9], projetos[10], projetos[11], projetos[12], projetos[13], projetos[14])
    //console.log(`projetos: ${projetos[6]}`)

    //Insere a nova linha com dados na resposta
    dados_final.push(item)

  })
  console.log(`itens: ${dados_final.length}`)
  console.log(`exemplo: ${dados_final[0]}`)
  return dados_final
}

function calcRateio(row) {

  const horas_total = row[16]
  const cni = horas_total * row[22] / 100
  const sesi = horas_total * row[23] / 100
  const senai = horas_total * row[24] / 100
  const iel = horas_total * row[25] / 100
  const horas_rateio = [cni, sesi, senai, iel]

  //console.log(`Horas totais: ${horas_rateio}`)

  return horas_rateio

}

function cutNames(name) {
  newName = name.split(" ")
  return (newName[0] + " " + newName[newName.length - 1])
}

function writeReport() {

  const rawData = getDadosCompletos()
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Relatorio')
  let reportDataToWrite = [] // Dados do relatorio

  let rateio = [0, 0, 0, 0] // inicia variavel de rateio

  rawData.forEach(function (item) {

    //inicia array que recebe cada card
    let rowToPush = []

    rowToPush.push(item[14]) // Projeto
    rowToPush.push(item[0]) // Card
    rowToPush.push(item[17]) // Descri??ao
    rowToPush.push(cutNames(item[12])) // autorizado por
    rowToPush.push(item[13]) // Autorizado em
    rowToPush.push(cutNames(item[2]))  // Atestado por
    rowToPush.push(item[11]) // Atestado em
    rowToPush.push(item[16]) // horas a faturar 

    //Insere a linha no array geral
    reportDataToWrite.push(rowToPush)

    //Calculo da soma das horas segundo rateio
    const rateioRow = calcRateio(item)
    rateio = rateioRow.map((item, i) => item + rateio[i])
    console.log(`RAteio: ${rateio}`)


  })

  //Escreve o numero de horas do rateio na planilha
  ws.getRange("G4").setValue(rateio[0]) // CNI
  ws.getRange("G5").setValue(rateio[1]) //SESI
  ws.getRange("G6").setValue(rateio[2]) // SENAI
  ws.getRange("G7").setValue(rateio[3]) // IEL

  // Escreve relatorio na planilha
  const rangeToClean = ws.getRange(12, 1, 1000, 50).clearContent()
  const range = ws.getRange(12, 1, reportDataToWrite.length, reportDataToWrite[1].length).setValues(reportDataToWrite)

  //console.log(`Final: ${reportDataToWrite}`)

}


