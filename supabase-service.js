// ─────────────────────────────────────────────────────────────────
// SUPABASE SERVICE — Skarpa Logística
// Centraliza o salvamento da planilha (Storage) e dados (Database)
// ─────────────────────────────────────────────────────────────────

// Inicializa o cliente Supabase utilizando a biblioteca global injetada via CDN
let supabase = null;

function getClient() {
    if (!supabase) {
        if (typeof window.supabase === 'undefined') {
            console.error("Erro: SDK do Supabase não foi carregado via CDN!");
            return null;
        }
        supabase = window.supabase.createClient(window.SUPABASE_URL, window.SUPABASE_ANON_KEY);
    }
    return supabase;
}

/**
 * Registra ou faz Login simples do usuário (Utilizando a tabela user_progress como identificador)
 */
export function sbLogin(email, password) {
    // Para simplificar a experiência do usuário e não precisar configurar e-mails de confirmação do Supabase Auth,
    // vamos usar um sistema de login baseado em dados na tabela 'user_progress'.
    // Caso queira usar Supabase Auth puro depois, o código está pronto para upgrade.
    return new Promise(async (resolve, reject) => {
        try {
            const client = getClient();
            if (!client) return reject("Cliente Supabase não iniciado");

            // Verifica se o usuário existe na nossa tabela
            const { data, error } = await client
                .from('user_progress')
                .select('*')
                .eq('email', email)
                .single();

            if (error && error.code !== 'PGRST116') { // PGRST116 = Nenhum resultado encontrado
                return reject("Erro ao conectar com o banco de dados.");
            }

            if (!data) {
                // Cria uma conta nova na hora se não existir
                const { error: insertError } = await client
                    .from('user_progress')
                    .insert([{ email: email, password: password, condominios: [], categories: {}, formatting: {} }]);

                if (insertError) return reject("Erro ao criar conta nova: " + insertError.message);
                
                resolve({ email, uid: email });
            } else {
                // Se existe, valida a senha simples
                if (data.password === password) {
                    resolve({ email: data.email, uid: data.email });
                } else {
                    reject("Senha incorreta.");
                }
            }
        } catch (e) {
            reject(e.message);
        }
    });
}

/**
 * Salva o arquivo .xlsx no Supabase Storage
 */
export async function sbSaveWorkbook(email, arrayBuffer) {
    const client = getClient();
    if (!client) return;

    // Converte ArrayBuffer em Blob
    const blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filePath = `${email}/planilha_atual.xlsx`;

    // Faz o upload no bucket público 'planilhas'
    const { error } = await client.storage
        .from('planilhas')
        .upload(filePath, blob, { upsert: true });

    if (error) {
        console.error("Erro ao salvar arquivo no Storage:", error.message);
    }
}

/**
 * Carrega a planilha salva no Supabase Storage
 */
export async function sbLoadWorkbook(email) {
    const client = getClient();
    if (!client) return null;

    const filePath = `${email}/planilha_atual.xlsx`;

    const { data, error } = await client.storage
        .from('planilhas')
        .download(filePath);

    if (error) {
        // Se der erro de arquivo inexistente, retornamos nulo
        return null;
    }

    return await data.arrayBuffer();
}

/**
 * Salva metadados (Categorias, Formatação e Condomínios) no Banco de Dados
 */
export async function sbSaveMetadata(email, { categories, formatting, condominios }) {
    const client = getClient();
    if (!client) return;

    const updateData = {};
    if (categories) updateData.categories = categories;
    if (formatting) updateData.formatting = formatting;
    if (condominios) updateData.condominios = condominios;

    const { error } = await client
        .from('user_progress')
        .update(updateData)
        .eq('email', email);

    if (error) {
        console.error("Erro ao salvar metadados:", error.message);
    }
}

/**
 * Carrega todos os metadados do Banco de Dados
 */
export async function sbLoadMetadata(email) {
    const client = getClient();
    if (!client) return null;

    const { data, error } = await client
        .from('user_progress')
        .select('categories, formatting, condominios')
        .eq('email', email)
        .single();

    if (error) {
        return { categories: {}, formatting: {}, condominios: [] };
    }

    return {
        categories: data.categories || {},
        formatting: data.formatting || {},
        condominios: data.condominios || []
    };
}
