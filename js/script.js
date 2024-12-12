const dropArea = document.getElementById('drop-area');
const fileInput = document.getElementById('file-input');

// Ao clicar na área, abre o seletor de arquivo
dropArea.addEventListener('click', () => fileInput.click());

// Para arrastar e soltar
dropArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropArea.style.borderColor = 'blue';
});

dropArea.addEventListener('dragleave', () => {
  dropArea.style.borderColor = '#ccc';
});

dropArea.addEventListener('drop', async (e) => {
  e.preventDefault();
  dropArea.style.borderColor = '#ccc';
  const file = e.dataTransfer.files[0];
  console.log('Drop');
  //if (file) handleFile(file);
  
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    function: handleFile,
    args: [file]
  });
});

// Para selecionar um arquivo
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  console.log('Change');
  //if (file) handleFile(file);
  
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    function: handleFile,
    args: [file]
  });
});

// Função para lidar com o arquivo Excel
async function handleFile(file) {
  const reader = new FileReader();

  // Quando o arquivo é lido, processa com SheetJS
  reader.onload = async (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Obtém a primeira planilha
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Converte os dados da planilha para JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    console.log('Dados da Planilha:', jsonData);
    insertInto(jsonData);

  };

  reader.readAsArrayBuffer(file);
}

function insertInto(jsonData) {
  const enterEvent = new KeyboardEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    keyCode: 13,
    bubbles: true,
    cancelable: true
  });
  
  const focusOutEvent = new FocusEvent('focusout', {
    bubbles: true,   
    cancelable: true 
  });
  
  jsonData.forEach((row, index) => {
    if (index === 0) return;
    if (isNaN(row[0])) return console.log(`Linha ${index} não contém um número válido. Conteúdo: ${row}`);
    
    const aluno = document.querySelector(`table[data-recordindex="${row[0]-1}"][id^="gridview"]`); //const aluno = document.querySelector('table[data-recordindex="'+ row[0] - 1 + '"][id^="gridview"]');
    if (!aluno) return console.log("Aluno com número" + row[0] + "não encontrado! \nConteúdo Linha: " + row);
    
    const dataToFill = row.slice(2); //Sobra só os 4 bimestres
    
    dataToFill.forEach((data, dataIndex) => {
      const cell = aluno.querySelector(`td:nth-child(${dataIndex + 3})`);
      if (!cell) return console.log(`Célula do bimestre ${dataIndex+1} não encontrada: ${cell}`);
      cell.dispatchEvent(enterEvent);

      const cellEdit = cell.querySelector('input');
      if (!cellEdit) return console.log(`Célula editável não encontrada para: ${cell}`);
      if (data == "") return;
      cellEdit.value = data;
      cellEdit.dispatchEvent(focusOutEvent);
    });
  });
}