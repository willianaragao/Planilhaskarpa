// ─────────────────────────────────────────────────────────────────
// APP.JS — Login com Supabase
// ─────────────────────────────────────────────────────────────────

import { sbLogin } from './supabase-service.js';

document.addEventListener('DOMContentLoaded', () => {
    const loginForm       = document.getElementById('loginForm');
    const usernameInput   = document.getElementById('username');
    const passwordInput   = document.getElementById('password');
    const togglePasswordBtn = document.getElementById('togglePassword');
    const errorMessage    = document.getElementById('errorMessage');
    const errorSpan       = errorMessage?.querySelector('span');
    const btnLogin        = document.getElementById('btnLogin');

    // Se já está logado no localStorage, podemos manter ou checar
    if (localStorage.getItem('isLoggedIn') === 'true' && localStorage.getItem('userEmail')) {
         window.location.href = 'dashboard.html';
    }

    // ── Toggle visibilidade da senha ─────────────────────────────
    togglePasswordBtn?.addEventListener('click', () => {
        const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        passwordInput.setAttribute('type', type);
        const icon = togglePasswordBtn.querySelector('i');
        icon.classList.toggle('fa-eye');
        icon.classList.toggle('fa-eye-slash');
    });

    // ── Submissão do formulário ──────────────────────────────────
    loginForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        errorMessage.style.display = 'none';
        btnLogin.classList.add('loading');

        const email    = usernameInput.value.trim();
        const password = passwordInput.value;

        try {
            const user = await sbLogin(email, password);
            
            localStorage.setItem('isLoggedIn', 'true');
            localStorage.setItem('userEmail', user.email);
            localStorage.setItem('username', user.email.split('@')[0]); // Usa o início do email como nome
            
            window.location.href = 'dashboard.html';
        } catch (err) {
            btnLogin.classList.remove('loading');
            if (errorSpan) errorSpan.textContent = typeof err === 'string' ? err : 'Usuário ou senha incorretos.';
            errorMessage.style.display = 'flex';
        }
    });

    // ── Limpa erro ao digitar ────────────────────────────────────
    [usernameInput, passwordInput].forEach(input => {
        input?.addEventListener('input', () => {
            errorMessage.style.display = 'none';
        });
    });
});
