import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";
import { getFirestore, doc, setDoc, getDoc } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";

// Tu Configuración[cite: 1]
const firebaseConfig = {
    apiKey: "AIzaSyCCPwgKILNrFt08pSJMxxH1W9Bc18JWBFE",
    authDomain: "pharma-task.firebaseapp.com",
    projectId: "pharma-task",
    storageBucket: "pharma-task.firebasestorage.app",
    messagingSenderId: "157821568633",
    appId: "1:157821568633:web:be0aa1626322119a1af6a8",
    measurementId: "G-1B6T04F8MH"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// Objeto Global de Autenticación
window.AuthUI = {
    async loginWithGoogle() {
        const provider = new GoogleAuthProvider();
        try {
            await signInWithPopup(auth, provider);
        } catch (e) { Swal.fire('Error', 'No se pudo conectar con Google', 'error'); }
    },
    async handleLogin() {
        const email = document.getElementById('login-email').value;
        const pass = document.getElementById('login-pass').value;
        if(!email || !pass) return Swal.fire('Aviso', 'Completa los datos', 'warning');
        try {
            await signInWithEmailAndPassword(auth, email, pass);
        } catch (e) {
            try { await createUserWithEmailAndPassword(auth, email, pass); }
            catch (err) { Swal.fire('Error', 'Credenciales incorrectas', 'error'); }
        }
    },
    logout() { signOut(auth).then(() => location.reload()); }
};

// Objeto Global de la App[cite: 1]
window.App = {
    tasks: [], clients: [], user: null,
    async init(user) {
        this.user = user;
        document.getElementById('auth-overlay').style.display = 'none';
        document.getElementById('user-display').innerText = user.email;
        await this.syncFromCloud();
        this.render();
    },
    async syncToCloud() {
        await setDoc(doc(db, "users", this.user.uid), { tasks: this.tasks, clients: this.clients });
        this.updateStats();
    },
    async syncFromCloud() {
        const snap = await getDoc(doc(db, "users", this.user.uid));
        if(snap.exists()) { this.tasks = snap.data().tasks || []; this.clients = snap.data().clients || []; }
        this.updateStats();
        this.updateClientUI();
    },
    handleExcel(e) {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const wb = XLSX.read(data, {type:'array'});
            const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
            this.clients = json.map(r => ({ id: (r.Codigo || "").toString(), name: (r.NombreFantasia || "").toString() }));
            this.syncToCloud();
            this.updateClientUI();
            Swal.fire('Éxito', 'Base de datos cargada', 'success');
        };
        reader.readAsArrayBuffer(e.target.files[0]);
    },
    addTask(t) { this.tasks.unshift({ id: Date.now(), status: 'pendiente', ...t }); this.syncToCloud(); this.render(); },
    deleteTask(id) { this.tasks = this.tasks.filter(t => t.id !== id); this.syncToCloud(); this.render(); },
    updateStatus(id, s) { const t = this.tasks.find(x => x.id === id); if(t){ t.status = s; this.syncToCloud(); this.render(); } },
    updateStats() {
        document.getElementById('stat-total').innerText = this.tasks.length;
        document.getElementById('stat-process').innerText = this.tasks.filter(t => t.status === 'proceso').length;
        document.getElementById('stat-done').innerText = this.tasks.filter(t => t.status === 'completada').length;
    },
    updateClientUI() {
        const el = document.getElementById('client-status');
        el.innerText = this.clients.length > 0 ? `✅ ${this.clients.length} farmacias` : '❌ Sin datos';
        el.style.color = this.clients.length > 0 ? 'var(--success)' : 'var(--danger)';
    },
    render() { UI.drawList(this.tasks); },
    toggleTheme() {
        const b = document.body;
        const theme = b.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        b.setAttribute('data-theme', theme);
    }
};

// Interfaz de Usuario
window.UI = {
    drawList(list) {
        const container = document.getElementById('taskListContent');
        container.innerHTML = list.map(t => `
            <div class="task-item">
                <div>
                    <strong>${t.pharmacy}</strong>
                    <p style="font-size:0.8rem; color:var(--text-muted)">${t.clientId} | ${t.desc}</p>
                </div>
                <div style="display:flex; gap:5px">
                    <button class="btn-icon" onclick="App.updateStatus(${t.id}, 'completada')">✅</button>
                    <button class="btn-icon" onclick="App.deleteTask(${t.id})">🗑️</button>
                </div>
            </div>
        `).join('');
    },
    async openModal() {
        const { value: form } = await Swal.fire({
            title: 'Nueva Tarea',
            html: `<input id="s-id" class="swal2-input" placeholder="Código"><input id="s-name" class="swal2-input" placeholder="Farmacia"><textarea id="s-desc" class="swal2-textarea" placeholder="Descripción"></textarea>`,
            preConfirm: () => ({ clientId: document.getElementById('s-id').value, pharmacy: document.getElementById('s-name').value, desc: document.getElementById('s-desc').value })
        });
        if(form) App.addTask(form);
    }
};

// Listener de Autenticación
onAuthStateChanged(auth, user => { if(user) App.init(user); else document.getElementById('auth-overlay').style.display='flex'; });