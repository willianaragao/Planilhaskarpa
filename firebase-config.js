// ─────────────────────────────────────────────────────────────────
// FIREBASE CONFIG — Skarpa Logística
// Preencha as credenciais abaixo com as do seu projeto Firebase.
// Como obter: https://console.firebase.google.com → seu projeto →
//             Configurações → Seus apps → SDK do Firebase
// ─────────────────────────────────────────────────────────────────

const FIREBASE_CONFIG = {
    apiKey:            "COLE_AQUI_SUA_apiKey",
    authDomain:        "COLE_AQUI_SEU_authDomain",
    projectId:         "COLE_AQUI_SEU_projectId",
    storageBucket:     "COLE_AQUI_SEU_storageBucket",
    messagingSenderId: "COLE_AQUI_SEU_messagingSenderId",
    appId:             "COLE_AQUI_SEU_appId"
};

// Exporta para os outros scripts usarem
window.FIREBASE_CONFIG = FIREBASE_CONFIG;
