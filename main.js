import './style.css';

const apiKeyInput = document.getElementById('apiKey');
const keyStatus = document.getElementById('keyStatus');
const saveApiKeyBtn = document.getElementById('saveApiKeyBtn');
const questionInput = document.getElementById('question');
const askBtn = document.getElementById('askBtn');
const statusLine = document.getElementById('statusLine');
const resultArea = document.getElementById('result-area');
const formulaText = document.getElementById('formulaText');
const explanationText = document.getElementById('explanationText');
const stepsList = document.getElementById('stepsList');
const exampleBlock = document.getElementById('exampleBlock');
const noteBlock = document.getElementById('noteBlock');
const copyBtn = document.getElementById('copyBtn');
const historyCard = document.getElementById('history-card');
const historyList = document.getElementById('historyList');

let history = JSON.parse(localStorage.getItem('gemini_history') || '[]');

let currentWorkbook = null;
let currentSheetName = null;
let currentData = null; // Array of objects
let originalFileName = 'modified_data.xlsx';

const fileInput = document.getElementById('fileInput');
const dataPreview = document.getElementById('dataPreview');
const spreadsheetWindow = document.getElementById('spreadsheetWindow');
const gridWrapper = document.getElementById('gridWrapper');
const expandOverlay = document.getElementById('expandOverlay');
const expandBtn = document.getElementById('expandBtn');
const downloadBtn = document.getElementById('downloadBtn');
const modeToggle = document.getElementById('modeToggle');
let jspreadsheetInstance = null;

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) {
    currentData = null;
    dataPreview.style.display = 'none';
    modeToggle.style.display = 'none';
    askBtn.textContent = 'Get formula';
    return;
  }
  originalFileName = file.name;
  
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, {type: 'array'});
      currentWorkbook = workbook;
      currentSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[currentSheetName];
      currentData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      
      renderPreview(currentData);
      modeToggle.style.display = 'block';
      document.querySelector('input[name="mode"][value="modify"]').checked = true;
      askBtn.textContent = 'Apply Changes';
    } catch (err) {
      alert('Error reading Excel file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

document.querySelectorAll('input[name="mode"]').forEach(radio => {
  radio.addEventListener('change', (e) => {
    askBtn.textContent = e.target.value === 'modify' ? 'Apply Changes' : 'Get formula';
  });
});

function renderPreview(data) {
  if (jspreadsheetInstance && typeof jspreadsheetInstance.destroy === 'function') {
    try { jspreadsheetInstance.destroy(); } catch(e) {}
  } else if (typeof jspreadsheet !== 'undefined' && jspreadsheet.destroy) {
    try { jspreadsheet.destroy(spreadsheetWindow); } catch(e) {}
  }
  
  spreadsheetWindow.innerHTML = '';
  if (!data || data.length === 0) {
    dataPreview.style.display = 'none';
    return;
  }
  
  const allKeys = new Set();
  data.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));
  const keys = Array.from(allKeys);
  
  const columns = keys.map(k => ({
    type: 'text',
    title: k,
    name: k,
    width: Math.max(110, Math.min(300, k.length * 11 + 25))
  }));
  
  // Transform objects into array of primitive values so Jspreadsheet renders properly regardless of syntax/version
  const rows = data.map(row => keys.map(k => (row[k] !== undefined && row[k] !== null) ? String(row[k]) : ''));
  
  jspreadsheetInstance = jspreadsheet(spreadsheetWindow, {
    data: rows,
    columns: columns,
    search: true,
    tableOverflow: true,
    tableWidth: "100%",
    tableHeight: data.length > 10 ? "300px" : "auto"
  });
  
  if (data.length > 10) {
    expandOverlay.style.display = 'flex';
    gridWrapper.style.maxHeight = '350px';
    expandBtn.textContent = 'Expand Rows';
    expandOverlay.style.background = 'linear-gradient(transparent, var(--mist) 80%)';
  } else {
    expandOverlay.style.display = 'none';
    gridWrapper.style.maxHeight = 'none';
  }
  
  dataPreview.style.display = 'block';
  downloadBtn.style.display = 'none';
}

downloadBtn.addEventListener('click', () => {
  if (!currentData) return;
  const newSheet = XLSX.utils.json_to_sheet(currentData);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, currentSheetName || "Sheet1");
  const outName = originalFileName.replace(/\.[^/.]+$/, "") + "_modified.xlsx";
  XLSX.writeFile(newWorkbook, outName);
});

function getStoredApiKey() {
  return (apiKeyInput.value || localStorage.getItem('gemini_api_key') || '').trim();
}

apiKeyInput.addEventListener('input', () => {
  const val = apiKeyInput.value.trim();
  const has = val.length > 10;
  keyStatus.classList.toggle('ok', has);
  keyStatus.innerHTML = has ? '<span class="dot"></span>Ready to save' : '<span class="dot"></span>Not set';
  if (has) {
    localStorage.setItem('gemini_api_key', val);
  } else {
    localStorage.removeItem('gemini_api_key');
  }
});

if (saveApiKeyBtn) {
  saveApiKeyBtn.addEventListener('click', () => {
    const val = apiKeyInput.value.trim();
    if (val.length > 10) {
      localStorage.setItem('gemini_api_key', val);
      keyStatus.classList.add('ok');
      keyStatus.innerHTML = '<span class="dot"></span>Key saved!';
      const origText = saveApiKeyBtn.textContent;
      saveApiKeyBtn.textContent = 'Saved!';
      saveApiKeyBtn.style.background = '#10B981';
      setTimeout(() => {
        saveApiKeyBtn.textContent = origText;
        saveApiKeyBtn.style.background = 'var(--excel)';
        const modal = document.getElementById('apiKeyModal');
        if (modal) modal.style.display = 'none';
      }, 1000);
    } else {
      localStorage.removeItem('gemini_api_key');
      keyStatus.classList.remove('ok');
      keyStatus.innerHTML = '<span class="dot" style="background:#EF4444"></span>Invalid API key';
      alert('Please enter a valid Gemini API key.');
    }
  });
}

document.addEventListener('DOMContentLoaded', () => {
  const savedKey = localStorage.getItem('gemini_api_key');
  if (savedKey) {
    apiKeyInput.value = savedKey;
    keyStatus.classList.add('ok');
    keyStatus.innerHTML = '<span class="dot"></span>Key set';
  } else {
    apiKeyInput.value = '';
    keyStatus.classList.remove('ok');
    keyStatus.innerHTML = '<span class="dot"></span>Not set';
  }
  renderHistory();
});

document.querySelectorAll('.chip').forEach(chip => {
  chip.addEventListener('click', () => {
    questionInput.value = chip.dataset.q;
    questionInput.focus();
  });
});

copyBtn.addEventListener('click', () => {
  navigator.clipboard.writeText(formulaText.textContent).then(() => {
    const original = copyBtn.textContent;
    copyBtn.textContent = 'Copied';
    setTimeout(() => copyBtn.textContent = original, 1400);
  });
});

const schema = {
  type: "OBJECT",
  properties: {
    formula: { type: "STRING" },
    explanation: { type: "STRING" },
    steps: { type: "ARRAY", items: { type: "STRING" } },
    example: { type: "STRING" },
    note: { type: "STRING" }
  },
  required: ["formula", "explanation", "steps"]
};

const systemPrompt = `You are an expert Microsoft Excel trainer. Given a plain-English description of a spreadsheet task, respond with:
- formula: the exact Excel formula that accomplishes the task, using generic cell/range references (A2, B2:B100, etc.) since you don't know the user's actual layout. Prefer modern Excel functions (XLOOKUP, IFS, UNIQUE, TEXTBEFORE, etc.) when they fit, but mention an older-Excel-compatible alternative in the note field if relevant.
- explanation: a 1-3 sentence plain-English explanation of how the formula works and why it solves the task.
- steps: an array of 3-6 short, concrete steps for how to enter/adapt this formula in their own sheet, referencing generic cell ranges the user should replace with their own.
- example: one short concrete worked example with sample input and output values.
- note: (optional) any caveat, edge case, or older-Excel-compatible alternative formula. Omit or leave empty if not needed.
Respond ONLY with the JSON object matching the schema. No markdown, no preamble.`;

const codeSchema = {
  type: "OBJECT",
  properties: {
    code: { type: "STRING" }
  },
  required: ["code"]
};

const codeSystemPrompt = `You are an expert JavaScript developer.
The user wants to modify an array of JavaScript objects representing Excel rows.
Write a valid JavaScript function named 'process(data)' that takes an array of objects 'data' and returns the modified array of objects.
Do not wrap it in markdown. Do not include any explanations. Just output the function itself inside a JSON property called 'code'.
If the request requires adding a column, modify each object to include the new property.
Keep it robust. Assume basic data types. Use standard JavaScript methods.`;

const MODEL_NAME = "gemini-3.5-flash";

async function modifyDataWithGemini(question, dataSample) {
  const apiKey = getStoredApiKey();
  const sampleText = JSON.stringify(dataSample, null, 2);
  const prompt = `Here is a sample of the data (first 3 rows):\n${sampleText}\n\nTask: ${question}`;
  
  const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: `${codeSystemPrompt}\n\n${prompt}` }] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: codeSchema,
        temperature: 0.1
      }
    })
  });

  if (!res.ok) {
    const errBody = await res.json().catch(() => null);
    const msg = errBody?.error?.message || `Request failed (${res.status})`;
    throw new Error(msg);
  }

  const resData = await res.json();
  const text = resData?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('No response returned from Gemini.');
  return JSON.parse(text).code;
}

async function askGemini(question) {
  const apiKey = getStoredApiKey();
  const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: `${systemPrompt}\n\nTask: ${question}` }] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: schema,
        temperature: 0.3
      }
    })
  });

  if (!res.ok) {
    const errBody = await res.json().catch(() => null);
    const msg = errBody?.error?.message || `Request failed (${res.status})`;
    throw new Error(msg);
  }

  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('No response returned from Gemini.');
  return JSON.parse(text);
}

function renderResult(result, question) {
  formulaText.textContent = result.formula || '';
  explanationText.textContent = result.explanation || '';

  stepsList.innerHTML = '';
  (result.steps || []).forEach(step => {
    const li = document.createElement('li');
    li.textContent = step;
    stepsList.appendChild(li);
  });

  if (result.example) {
    exampleBlock.style.display = 'block';
    exampleBlock.innerHTML = `<strong>Example:</strong> ${result.example}`;
  } else {
    exampleBlock.style.display = 'none';
  }

  if (result.note) {
    noteBlock.style.display = 'block';
    noteBlock.textContent = result.note;
  } else {
    noteBlock.style.display = 'none';
  }

  resultArea.style.display = 'block';
  resultArea.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  history.unshift({ question, formula: result.formula });
  history = history.slice(0, 8);
  localStorage.setItem('gemini_history', JSON.stringify(history));
  renderHistory();
}

function renderHistory() {
  if (history.length === 0) { historyCard.style.display = 'none'; return; }
  historyCard.style.display = 'block';
  historyList.innerHTML = '';
  history.forEach(item => {
    const div = document.createElement('div');
    div.className = 'history-item';
    div.innerHTML = `<div class="history-q">${item.question}</div><div class="history-f">${item.formula}</div>`;
    div.addEventListener('click', () => {
      questionInput.value = item.question;
      questionInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
    historyList.appendChild(div);
  });
}

async function handleAsk() {
  const apiKey = getStoredApiKey();
  const question = questionInput.value.trim();

  statusLine.classList.remove('error');

  if (!apiKey) {
    statusLine.textContent = 'Add your Gemini API key above first.';
    statusLine.classList.add('error');
    apiKeyInput.focus();
    return;
  }
  if (!question) {
    statusLine.textContent = 'Describe what you want the formula to do.';
    statusLine.classList.add('error');
    questionInput.focus();
    return;
  }

  askBtn.disabled = true;
  askBtn.innerHTML = '<span class="spinner"></span>Thinking…';
  statusLine.textContent = '';

  try {
    const selectedMode = document.querySelector('input[name="mode"]:checked')?.value || 'formula';
    if (currentData && selectedMode === 'modify') {
      statusLine.textContent = 'Writing transformation code...';
      const dataSample = currentData.slice(0, 3);
      const jsCode = await modifyDataWithGemini(question, dataSample);
      
      statusLine.textContent = 'Applying transformation...';
      const funcStr = jsCode.trim();
      let processFn;
      try {
        processFn = new Function('return ' + funcStr)();
        if (typeof processFn !== 'function') throw new Error('Returned code is not a function.');
      } catch (e) {
        processFn = new Function('data', funcStr + '\nreturn process(data);');
      }
      
      const modifiedData = processFn(JSON.parse(JSON.stringify(currentData)));
      currentData = modifiedData;
      renderPreview(currentData);
      
      statusLine.textContent = 'Data successfully modified!';
      statusLine.classList.remove('error');
      downloadBtn.style.display = 'inline-block';
      resultArea.style.display = 'none';
    } else {
      const result = await askGemini(question);
      renderResult(result, question);
      statusLine.textContent = '';
    }
  } catch (err) {
    statusLine.textContent = err.message || 'Something went wrong. Check your API key and try again.';
    statusLine.classList.add('error');
  } finally {
    askBtn.disabled = false;
    const finalMode = document.querySelector('input[name="mode"]:checked')?.value || 'formula';
    askBtn.textContent = (currentData && finalMode === 'modify') ? 'Apply Changes' : 'Get formula';
  }
}

askBtn.addEventListener('click', handleAsk);
questionInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) handleAsk();
});

const cheatSheetBtn = document.getElementById('cheatSheetBtn');
const cheatSheetModal = document.getElementById('cheatSheetModal');
const closeCheatSheet = document.getElementById('closeCheatSheet');

if (cheatSheetBtn && cheatSheetModal && closeCheatSheet) {
  cheatSheetBtn.addEventListener('click', () => {
    cheatSheetModal.style.display = 'flex';
  });

  closeCheatSheet.addEventListener('click', () => {
    cheatSheetModal.style.display = 'none';
  });

  cheatSheetModal.addEventListener('click', (e) => {
    if (e.target === cheatSheetModal) {
      cheatSheetModal.style.display = 'none';
    }
  });
}

const apiKeyBtn = document.getElementById('apiKeyBtn');
const apiKeyModal = document.getElementById('apiKeyModal');
const closeApiKeyModal = document.getElementById('closeApiKeyModal');

if (apiKeyBtn && apiKeyModal && closeApiKeyModal) {
  apiKeyBtn.addEventListener('click', () => {
    apiKeyModal.style.display = 'flex';
  });

  closeApiKeyModal.addEventListener('click', () => {
    apiKeyModal.style.display = 'none';
  });

  apiKeyModal.addEventListener('click', (e) => {
    if (e.target === apiKeyModal) {
      apiKeyModal.style.display = 'none';
    }
  });
}

if (expandBtn && gridWrapper) {
  expandBtn.addEventListener('click', () => {
    if (gridWrapper.style.maxHeight === 'none') {
      gridWrapper.style.maxHeight = '350px';
      expandBtn.textContent = 'Expand Rows';
      expandOverlay.style.background = 'linear-gradient(transparent, var(--mist) 80%)';
    } else {
      gridWrapper.style.maxHeight = 'none';
      expandBtn.textContent = 'Collapse Rows';
      expandOverlay.style.background = 'transparent';
    }
  });
}
