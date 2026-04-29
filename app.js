document.addEventListener('DOMContentLoaded', () => {
    const loginForm = document.getElementById('loginForm');
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const togglePasswordBtn = document.getElementById('togglePassword');
    const errorMessage = document.getElementById('errorMessage');
    const btnLogin = document.getElementById('btnLogin');

    // Hardcoded credentials for demonstration
    const VALID_USER = 'admin';
    const VALID_PASS = '123456';

    // Toggle Password Visibility
    togglePasswordBtn.addEventListener('click', () => {
        const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        passwordInput.setAttribute('type', type);
        
        // Toggle icon
        const icon = togglePasswordBtn.querySelector('i');
        if (type === 'text') {
            icon.classList.remove('fa-eye');
            icon.classList.add('fa-eye-slash');
        } else {
            icon.classList.remove('fa-eye-slash');
            icon.classList.add('fa-eye');
        }
    });

    // Form Submission
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        
        // Reset state
        errorMessage.style.display = 'none';
        btnLogin.classList.add('loading');

        const username = usernameInput.value.trim();
        const password = passwordInput.value;

        // Simulate API Call
        setTimeout(() => {
            btnLogin.classList.remove('loading');

            if (username === VALID_USER && password === VALID_PASS) {
                // Success
                localStorage.setItem('isLoggedIn', 'true');
                localStorage.setItem('username', username);
                
                // Redirect to dashboard (to be created)
                window.location.href = 'dashboard.html';
            } else {
                // Failure
                errorMessage.style.display = 'flex';
                // Shake animation trigger
                errorMessage.style.animation = 'none';
                errorMessage.offsetHeight; // trigger reflow
                errorMessage.style.animation = null;
            }
        }, 1500); // 1.5s delay for realistic feel
    });

    // Clear error message on input
    [usernameInput, passwordInput].forEach(input => {
        input.addEventListener('input', () => {
            errorMessage.style.display = 'none';
        });
    });
});
