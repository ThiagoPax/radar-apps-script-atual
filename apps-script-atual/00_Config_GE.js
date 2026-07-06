// =============================================================================
// 00_Config_GE.js — Configurações centrais do Radar Gazeta Esportiva
// =============================================================================
// Sistema exclusivo do programa Gazeta Esportiva (18:00–18:59, seg–sex).
// A lógica do Jogo Aberto foi removida em 07/2026 (backup: Radar_Backup_pre_migracao_2026-07-04).

const CFG = {
    SPREADSHEET_ID: "11L0-jgTgMA3JOyuUfowsOzZ0svnC3AegNgekTA62LmI",
    ARQUIVO_NOME: "Radar_Gazeta_Esportiva",
    TIMEZONE: "America/Sao_Paulo",

    // -------------------------------------------------------------------------
    // Abas — nomes curtos (pós-migração). As chaves GE_OUT_* são mantidas
    // porque o motor de predição/painel/webhook as referencia.
    // -------------------------------------------------------------------------
    SHEET_HOJE: "Hoje",
    INPUT_SHEET: "Input",
    GE_INPUT_SHEET: "Input",
    INPUT_SHEET_NAME: "Input",
    GE_OUT_HISTORICO: "Histórico",
    HIST_SHEET_NAME: "Histórico",
    SHEET_RESUMO: "Resumo",
    SHEET_SEMANA: "Semana",
    SHEET_MES: "Mês",
    GE_OUT_TIME: "Minuto",
    GE_OUT_MEDIA_TEMA: "Tema Minuto",
    SHEET_DIA_SEMANA: "Dia Semana",
    SHEET_COMENTARISTAS: "Comentaristas",
    GE_OUT_TOP: "Termos",
    GE_OUT_MACRO: "Macrotema",
    GE_OUT_SUB: "Subtema",
    SHEET_BREAKS: "Breaks",
    SHEET_CLUBES: "Clubes",
    SHEET_EMAIL: "Email",
    SHEET_PDF: "PDF",
    SHEET_CONFIG: "Config",
    SHEET_LOGS: "Logs",
    SHEET_BASE: "Base",
    SHEET_BACKTEST: "Backtest",

    // Ordem operacional das abas (as demais vão para o fim; técnicas ocultas)
    ORDEM_ABAS: ["Hoje", "Input", "Histórico", "Resumo", "Semana", "Mês",
        "Minuto", "Tema Minuto", "Dia Semana", "Comentaristas", "Termos",
        "Macrotema", "Subtema", "Breaks", "Clubes", "Email", "PDF",
        "Config", "Logs", "Base", "Backtest"
    ],
    ABAS_OCULTAS: ["Email", "PDF", "Config", "Logs", "Base", "Backtest",
        "Fatores", "Backtest_Shadow", "Backtest_CP"],

    // -------------------------------------------------------------------------
    // Horários do programa e das rotinas
    // -------------------------------------------------------------------------
    GE_START_HOUR: 18,
    GE_START_MINUTE: 0,
    GE_END_HOUR: 18,
    GE_END_MINUTE: 59,
    ROTINA_MANHA_HOUR: 6,          // 06:00 — calendário, médias, breaks, preparar Hoje
    IMPORT_TENTATIVAS: [
        { hour: 19, minute: 0 },   // 1ª tentativa de localizar transcrição
        { hour: 19, minute: 5 },   // 2ª tentativa
        { hour: 19, minute: 10 }   // 3ª tentativa — principal
    ],
    IMPORT_ALERTA_HOUR: 19,        // 19:15 — alerta se ainda não houver transcrição
    IMPORT_ALERTA_MINUTE: 15,
    EMAIL_DASHBOARD_HOUR: 21,      // 21:00 — email dashboard do dia
    EMAIL_DASHBOARD_MINUTE: 0,
    // Compatibilidade com módulos existentes (importação Speech)
    IMPORT_TRIGGER_HOUR: 19,
    IMPORT_TRIGGER_MINUTE: 5,
    IMPORT_RETRY_TRIGGER_HOUR: 19,
    IMPORT_RETRY_TRIGGER_MINUTE: 10,
    GE_AUTO_RUN_HOUR: 21,
    GE_AUTO_RUN_MINUTE: 0,

    // -------------------------------------------------------------------------
    // YouTube / transcrição
    // -------------------------------------------------------------------------
    GE_CHANNEL_ID: "UCcndglQZmXnYjgE_Grqt5yg",
    GE_UPLOADS_PLAYLIST_ID: "UUcndglQZmXnYjgE_Grqt5yg",
    GE_TITLE_PATTERN: "Gazeta Esportiva | AO VIVO",
    GE_APRESENTADORAS: ["MICHELLE", "MARINA"],
    TRANSCRIPTION_SERVER: "",
    SPEECH_SOURCE_MODE: "DRIVE",
    SPEECH_DRIVE_FOLDER_ID: "1yYZf8oyRcBfEwUy_9_vuZIJxExYwEpGz",
    SPEECH_WEBHOOK_ENABLED: true,
    SPEECH_DRIVE_FILE_PREFIX: "transcricao",
    // false: as tentativas 19:00/19:05/19:10 não alertam; o alerta único
    // é responsabilidade do handler geAlertaTranscricaoGE às 19:15.
    SPEECH_ALERT_ON_MISSING: false,
    SPEECH_PROCESSED_PREFIX: "speech_processed_",

    // -------------------------------------------------------------------------
    // Email
    // -------------------------------------------------------------------------
    EMAIL_DESTINO: "tssouza@tssouza.com",
    EMAIL_ENABLED: true,

    // -------------------------------------------------------------------------
    // LLM / APIs externas
    // -------------------------------------------------------------------------
    OPENAI_ENDPOINT: "https://api.openai.com/v1/chat/completions",
    OPENAI_MODEL: "gpt-4o",
    LLM_ENABLED: true,
    LLM_BATCH_SIZE: 30,
    LLM_TEMPERATURE: 0,
    LLM_MAX_TOKENS: 4096,
    GEMINI_ENDPOINT: "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent",
    GEMINI_MODEL: "gemini-2.0-flash-lite",
    API_FOOTBALL_ENDPOINT: "https://v3.football.api-sports.io",

    // Os 4 grandes paulistas são o foco editorial; demais IDs ficam para
    // detecção de adversários/clássicos.
    CLUBES_FOCO: ["CORINTHIANS", "PALMEIRAS", "SÃO PAULO", "SANTOS"],
    TEAM_IDS: {
        CORINTHIANS: 131,
        PALMEIRAS: 121,
        "SÃO PAULO": 126,
        SANTOS: 128,
        FLAMENGO: 127,
        FLUMINENSE: 124,
        VASCO: 133,
        BOTAFOGO: 120,
        INTERNACIONAL: 119,
        GRÊMIO: 130,
        "ATLÉTICO-MG": 1062,
        CRUZEIRO: 129,
        BAHIA: 118,
        FORTALEZA: 1193,
        "ATHLETICO-PR": 134,
        BRAGANTINO: 1154
    },

    // -------------------------------------------------------------------------
    // Processamento de transcrição / temas
    // -------------------------------------------------------------------------
    TOP_TERMS_PER_MINUTE: 5,
    THEME_CARRY_MINUTES: 10,
    CARRY_MIN_CONFIDENCE: 0.55,
    MERCHAN_EXPLICIT_THRESHOLD: 8,
    MERCHAN_IMPLICIT_THRESHOLD: 12,
    MERCHAN_CARRY_MINUTES: 2,
    SMOOTH_WINDOW: 3,

    // Taxonomia do Gazeta Esportiva. Cobre o cotidiano dos clubes paulistas
    // E eventos excepcionais (Copa do Mundo 2026 em andamento).
    FIXED_THEMES: [
        // Clubes foco
        "CORINTHIANS", "PALMEIRAS", "SÃO PAULO", "SANTOS",
        // Copa do Mundo / seleções / personagens
        "COPA DO MUNDO", "SELEÇÃO BRASILEIRA", "ARGENTINA", "FRANÇA",
        "INGLATERRA", "NORUEGA", "MÉXICO", "FIFA", "NEYMAR", "ANCELOTTI",
        // Transversais
        "ARBITRAGEM", "POLÊMICA", "BASTIDOR",
        // Demais clubes brasileiros
        "FLAMENGO", "FLUMINENSE", "VASCO", "BOTAFOGO", "INTERNACIONAL",
        "GRÊMIO", "ATLÉTICO-MG", "CRUZEIRO", "BAHIA", "VITÓRIA", "FORTALEZA",
        "CEARÁ", "ATHLETICO-PR", "CORITIBA", "GOIÁS", "SPORT", "BRAGANTINO",
        // Internacional / outros esportes
        "FUTEBOL INTERNACIONAL", "F1", "NFL", "NBA", "FUTSAL", "VÔLEI", "MMA",
        "OUTROS_ESPORTES",
        // Operacionais
        "MERCHAN", "GERAL"
    ],
    SUBTEMAS_SUGERIDOS: ["JOGO DO DIA", "JOGO", "CLASSIFICAÇÃO", "OITAVAS",
        "QUARTAS", "FORMATO DA COPA", "CONVOCAÇÃO", "POLÊMICA", "BASTIDOR",
        "ARBITRAGEM", "REBAIXAMENTO", "RESENHA", "MERCADO", "OUTROS"
    ],

    // -------------------------------------------------------------------------
    // Contexto ativo — muda a régua de comparação (aba Config)
    // -------------------------------------------------------------------------
    CONTEXTOS_VALIDOS: ["NORMAL", "COPA DO MUNDO", "SEMANA DE CLÁSSICO",
        "SEM CLUBES PAULISTAS", "FERIADO", "GRADE ATÍPICA",
        "CONCORRÊNCIA FORTE", "QUEDA ESTRUTURAL"
    ],
    CONTEXTO_JANELA_CURTA: 3,   // últimos N programas
    CONTEXTO_JANELA_MEDIA: 5
};

// Cache em memória por execução (lexicon de entidades e stopwords)
const MEMORY_CACHE = {
    entityLexicon: null,
    stopwords: null
};
