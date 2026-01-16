// State
let allNames = []; // The pool of names
let remainingNames = []; // Names eligible to win
let winnersHistory = []; // List of winners
let isRolling = false;

// DOM Elements
const excelInput = document.getElementById('excelInput');
const fileInfo = document.getElementById('fileInfo');
const bgInput = document.getElementById('bgInput');
const delayInput = document.getElementById('delayInput');
const delayValue = document.getElementById('delayValue');
const resetBtn = document.getElementById('resetBtn');
const drawBtns = document.querySelectorAll('.draw-btn[data-count]');
const rollingContainer = document.getElementById('rollingContainer');
const readyText = document.getElementById('readyText');
const rollingDisplay = document.getElementById('rollingDisplay');
// Grand Modal Elements
const grandRevealModal = document.getElementById('grandRevealModal');
const grandWinnersList = document.getElementById('grandWinnersList');
const closeGrandBtn = document.getElementById('closeGrandBtn');

const winnersGrid = document.getElementById('winnersGrid');
const historyTableBody = document.getElementById('historyBody');
const totalWinners = document.getElementById('totalWinners');
const exportBtn = document.getElementById('exportBtn');
const customDrawInput = document.getElementById('customDrawInput');
const customDrawBtn = document.getElementById('customDrawBtn');

// Event Listeners
excelInput.addEventListener('change', handleExcelUpload);
bgInput.addEventListener('change', handleBgUpload);
delayInput.addEventListener('input', (e) => delayValue.innerText = e.target.value);
drawBtns.forEach(btn => btn.addEventListener('click', () => startDraw(parseInt(btn.dataset.count))));
resetBtn.addEventListener('click', resetGame);
exportBtn.addEventListener('click', exportHistory);
closeGrandBtn.addEventListener('click', () => {
    grandRevealModal.classList.add('hidden');
    fireConfetti(); // Extra confetti on close
});
customDrawBtn.addEventListener('click', () => {
    const val = parseInt(customDrawInput.value);
    if (val > 0) startDraw(val);
    else alert("Please enter a valid number");
});

// Initialize
function init() {
    // Check if XLSX is loaded
    if (typeof XLSX === 'undefined') {
        alert('Library xlsx not loaded. Please ensure internet connection or local file exists.');
    }
}
init();

// --- Excel Handling ---
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    fileInfo.innerText = "Loading...";

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assume first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        // Convert to JSON (array of arrays)
        const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // specific format: Cols A-F combined. row[0] to row[5]
        const processedNames = [];

        // Skip header if needed? usually row 1. Let's try to detect or just parse all.
        // User said "1 row counts as 1 person".
        rows.forEach((row, index) => {
            // Filter out empty rows
            if (row.length === 0) return;

            // Combine cols A-F (indices 0-5)
            // Filter undefined/null values
            const parts = row.slice(0, 6).filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== "");

            if (parts.length > 0) {
                const combined = parts.join('_'); // Separator changed to underscore
                processedNames.push(combined);
            }
        });

        if (processedNames.length > 0) {
            allNames = processedNames;
            remainingNames = [...allNames];
            fileInfo.innerText = `Loaded ${remainingNames.length} entries. Ready via Excel.`;
            fileInfo.style.color = '#4ade80'; // Green
        } else {
            fileInfo.innerText = "No valid data found.";
            fileInfo.style.color = '#ff6b6b';
        }
    };
    reader.readAsArrayBuffer(file);
}

// --- Background Handling ---
function handleBgUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const url = URL.createObjectURL(file);
    document.body.style.backgroundImage = `url(${url})`;
    document.body.style.backgroundSize = "cover";
    document.body.style.backgroundPosition = "center";
}

// --- Game Logic ---
function startDraw(count) {
    // Validation
    if (isNaN(count) || count <= 0) {
        console.error("Invalid count:", count);
        return;
    }

    if (remainingNames.length === 0) {
        alert("No names loaded or all names have won!");
        return;
    }

    if (isRolling) return;

    // Validate count vs remaining
    if (count > remainingNames.length) {
        alert(`Not enough names remaining! Only ${remainingNames.length} left.`);
        count = remainingNames.length;
    }

    isRolling = true;
    setControlsState(false);

    // UI Setup
    rollingDisplay.classList.remove('hidden');
    winnersGrid.classList.add('hidden');
    winnersGrid.innerHTML = ''; // Clear previous
    readyText.classList.add('hidden'); // Hide "Ready to Roll"
    rollingDisplay.innerText = "Initializing...";

    // Init Audio Context
    AudioManager.init();

    // Animation Variables
    const duration = parseFloat(delayInput.value) * 1000;
    const intervalTime = 50; // ms per name switch
    let elapsedTime = 0;

    // The "Rolling" Effect
    const rollInterval = setInterval(() => {
        // Show a random name from the remaining pool
        const randomIdx = Math.floor(Math.random() * remainingNames.length);
        rollingDisplay.innerText = remainingNames[randomIdx];

        AudioManager.playTick(); // Tick sound

        elapsedTime += intervalTime;

        if (elapsedTime >= duration) {
            clearInterval(rollInterval);
            finalizeDraw(count);
        }
    }, intervalTime);
}

function finalizeDraw(count) {
    // Layout Logic: ALWAYS SINGLE COLUMN now (User Revert)
    winnersGrid.classList.remove('multi-col');
    winnersGrid.classList.add('single-col');

    // Select Winners
    const currentWinners = [];
    for (let i = 0; i < count; i++) {
        if (remainingNames.length === 0) break;

        const winningIdx = Math.floor(Math.random() * remainingNames.length);
        const winner = remainingNames[winningIdx];

        // Double check against history (Safety net)
        if (winnersHistory.includes(winner)) {
            console.warn("Duplicate detected and skipped (Logic Safety):", winner);
            remainingNames.splice(winningIdx, 1); // Remove and retry
            i--;
            continue;
        }

        currentWinners.push(winner);
        remainingNames.splice(winningIdx, 1); // Remove from pool
    }

    // Update State
    winnersHistory.push(...currentWinners);

    // Update UI
    rollingDisplay.classList.add('hidden');
    winnersGrid.classList.remove('hidden');

    // Render Winners Cards
    currentWinners.forEach((name, index) => {
        const card = document.createElement('div');
        card.className = 'winner-card';
        card.innerText = name;
        // Stagger animation
        card.style.animationDelay = `${index * 0.1}s`;
        winnersGrid.appendChild(card);
    });

    // Update History
    updateHistoryTable();
    totalWinners.innerText = `Total: ${winnersHistory.length}`;

    // Reset State
    isRolling = false;
    setControlsState(true);

    // Trigger Effects
    // Trigger Effects
    fireConfetti();
    AudioManager.playWin(); // Win sound

    // Grand Modal Check (1 or 3 winners)
    if (count <= 3) {
        showGrandReveal(currentWinners);
    }
}

function showGrandReveal(winners) {
    grandWinnersList.innerHTML = '';
    winners.forEach(name => {
        const item = document.createElement('div');
        item.className = 'grand-winner-item';
        item.innerText = name;
        grandWinnersList.appendChild(item);
    });
    grandRevealModal.classList.remove('hidden');
}

function updateHistoryTable() {
    historyTableBody.innerHTML = '';
    // Reverse order to show latest top? Or 1-N? User asked "order 1 2 3 4".
    // "มีตารางแสดงชือที่ออกไปแล้วเรียงตามลำดับการสุ่ม 1 2 3 4"
    // So 1st winner is #1.
    winnersHistory.forEach((name, idx) => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${idx + 1}</td><td>${name}</td>`;
        historyTableBody.appendChild(row);
    });
}

function setControlsState(enabled) {
    const panels = document.querySelectorAll('button, input');
    panels.forEach(p => {
        if (enabled) p.parentElement.classList.remove('dimmed');
        else p.parentElement.classList.add('dimmed');
        p.disabled = !enabled;
    });
}

function resetGame() {
    if (!confirm("Are you sure you want to reset all history?")) return;
    remainingNames = [...allNames]; // Reset pool
    winnersHistory = [];
    updateHistoryTable();
    winnersGrid.innerHTML = '';
    rollingDisplay.innerText = "";
    rollingDisplay.classList.add('hidden');
    readyText.classList.remove('hidden'); // Show ready text again
    totalWinners.innerText = "Total: 0";
    alert("Game Reset!");
}

function exportHistory() {
    if (winnersHistory.length === 0) {
        alert("No history to export.");
        return;
    }

    const ws_data = [["No.", "Name"]];
    winnersHistory.forEach((name, idx) => {
        ws_data.push([idx + 1, name]);
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Winners");

    const filename = `LuckyDraw_Winners_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, filename);
}

// --- Confetti Effect (Simple Particle System) ---
function fireConfetti() {
    // Simple DOM based confetti
    const colors = ['#ff007a', '#7a00ff', '#00ff7a', '#ffae00', '#00eaff'];
    const container = document.body;

    for (let i = 0; i < 100; i++) {
        const confetto = document.createElement('div');
        confetto.style.position = 'absolute';
        confetto.style.width = '10px';
        confetto.style.height = '10px';
        confetto.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
        confetto.style.left = Math.random() * 100 + 'vw';
        confetto.style.top = '-10px';
        confetto.style.opacity = '1';
        confetto.style.zIndex = '9999';
        confetto.style.transform = `rotate(${Math.random() * 360}deg)`;

        // Random physics
        const duration = Math.random() * 2 + 2; // 2-4s

        confetto.style.transition = `top ${duration}s ease-out, opacity ${duration}s ease-in`;

        container.appendChild(confetto);

        // Animation
        requestAnimationFrame(() => {
            confetto.style.top = '110vh';
            confetto.style.opacity = '0';
        });

        // Cleanup
        setTimeout(() => {
            confetto.remove();
        }, duration * 1000);
    }
}

// --- Audio System (Synthesizer) ---
const AudioManager = {
    ctx: null,

    init: function () {
        if (!this.ctx) {
            const AudioContext = window.AudioContext || window.webkitAudioContext;
            this.ctx = new AudioContext();
        }
        if (this.ctx.state === 'suspended') {
            this.ctx.resume();
        }
    },

    playTick: function () {
        if (!this.ctx) return;
        const osc = this.ctx.createOscillator();
        const gain = this.ctx.createGain();

        osc.type = 'square';
        osc.frequency.setValueAtTime(800, this.ctx.currentTime);
        osc.frequency.exponentialRampToValueAtTime(100, this.ctx.currentTime + 0.05);

        gain.gain.setValueAtTime(0.05, this.ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, this.ctx.currentTime + 0.05);

        osc.connect(gain);
        gain.connect(this.ctx.destination);

        osc.start();
        osc.stop(this.ctx.currentTime + 0.05);
    },

    playWin: function () {
        if (!this.ctx) return;

        // Major Chord Fanfare: C4 E4 G4 C5
        const notes = [261.63, 329.63, 392.00, 523.25];
        const now = this.ctx.currentTime;

        notes.forEach((freq, i) => {
            const osc = this.ctx.createOscillator();
            const gain = this.ctx.createGain();

            osc.type = 'triangle';
            osc.frequency.value = freq;

            // Stagger start slightly
            const startTime = now + (i * 0.1);

            gain.gain.setValueAtTime(0, startTime);
            gain.gain.linearRampToValueAtTime(0.2, startTime + 0.1);
            gain.gain.exponentialRampToValueAtTime(0.001, startTime + 2);

            osc.connect(gain);
            gain.connect(this.ctx.destination);

            osc.start(startTime);
            osc.stop(startTime + 2);
        });
    }
};
