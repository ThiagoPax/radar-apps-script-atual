// =============================================================================
// 02_NLP.js — NLP: stopwords, entity lexicon, tokenizers, top terms, Vitória
// =============================================================================

// --- Stopwords ---

const STOPWORDS_STORAGE_KEY = "USER_STOPWORDS";

function getStopwords_() {
    return MEMORY_CACHE.stopwords || (MEMORY_CACHE.stopwords = buildStopwordsFromTop20_()), MEMORY_CACHE.stopwords
}

function loadPersistedStopwords_() {
    const e = PropertiesService.getScriptProperties().getProperty("USER_STOPWORDS");
    if (!e) return new Set;
    try {
        const t = JSON.parse(e);
        return new Set(t)
    } catch (e) {
        return new Set
    }
}

function savePersistedStopwords_(e) {
    const t = [...e];
    PropertiesService.getScriptProperties().setProperty("USER_STOPWORDS", JSON.stringify(t))
}

function addToPersistedStopwords_(e) {
    const t = loadPersistedStopwords_();
    for (const o of e) {
        const e = normalizeText_(o);
        e && t.add(e)
    }
    return savePersistedStopwords_(t), t
}

function buildBaseStopwords_() {
    const e = ["a", "ao", "aos", "as", "\xe0", "\xe0s", "aqui", "ali", "ai", "ae", "agora", "ainda", "alguem", "algum", "alguma", "ano", "anos", "antes", "apenas", "apos", "assim", "ate", "bem", "boa", "bom", "cada", "cai", "cair", "cara", "casa", "com", "como", "contra", "coisa", "coisas", "da", "das", "de", "dele", "dela", "deles", "delas", "demais", "depois", "desde", "dessa", "desse", "deste", "disso", "do", "dos", "duas", "e", "eh", "ela", "elas", "ele", "eles", "em", "entre", "entao", "essa", "esse", "esta", "estao", "estava", "este", "eu", "faz", "fazer", "fez", "fica", "ficar", "foi", "for", "fora", "fui", "gente", "grande", "ha", "hoje", "hora", "isso", "ja", "la", "mais", "mas", "me", "mesmo", "meu", "minha", "muito", "na", "nas", "nem", "no", "nos", "nossa", "nosso", "nossos", "num", "numa", "o", "os", "ou", "para", "pelo", "pela", "pelos", "pelas", "pois", "por", "porque", "pra", "pro", "pros", "pras", "quando", "que", "quem", "se", "sem", "sera", "ser", "sim", "so", "sobre", "ta", "tava", "tem", "tendo", "tenho", "ter", "teve", "to", "tudo", "um", "uma", "voce", "voces", "vou", "vamos", "vai", "t\xe1", "t\xf4", "tipo", "assim", "entendeu", "sabe", "olha", "veja", "fala", "falando", "dizer", "diz", "daqui", "pouco", "mesma", "mesmas", "mesmos", "tambem", "tamb\xe9m", "muita", "muitas", "muitos", "aplausos", "risos", "inaudivel", "inaud\xedvel", "musica", "m\xfasica"],
        t = new Set;
    for (const o of e) t.add(normalizeText_(o));
    return t
}

function buildStopwordsFromTop20_() {
    const e = buildBaseStopwords_(),
        t = loadPersistedStopwords_();
    for (const o of t) e.add(o);
    const o = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_TOP);
    if (o) {
        const t = o.getLastRow();
        if (t >= 2) {
            const a = o.getRange(2, 1, t - 1, 4).getValues();
            for (const t of a) {
                const o = normalizeText_(t[0] || ""),
                    a = !0 === t[3];
                o && a && e.add(o)
            }
        }
    }
    return e
}

function clearUserStopwords() {
    PropertiesService.getScriptProperties().deleteProperty("USER_STOPWORDS"), MEMORY_CACHE.stopwords = null, SpreadsheetApp.getUi().alert("Stopwords personalizadas removidas. As marcações serão resetadas na próxima execução.")
}

// --- Entity Lexicon ---

function getEntityLexicon_() {
    return MEMORY_CACHE.entityLexicon || (MEMORY_CACHE.entityLexicon = buildEntityLexicon_()), MEMORY_CACHE.entityLexicon
}

function buildEntityLexicon_() {
    const e = {};
    return addAliases_(e, "FLAMENGO", ["flamengo", "mengao", "meng\xe3o", "rubro negro", "rubronegro", "rubro-negro"]), addAliases_(e, "PALMEIRAS", ["palmeiras", "verdao", "verd\xe3o", "abel ferreira", "flaco lopez", "flaco lopes", "leila pereira"]), addAliases_(e, "CORINTHIANS", ["corinthians", "timao", "tim\xe3o", "fiel", "itaquera", "arena do timao", "dourival", "dorival", "corinthiano"]), addAliases_(e, "S\xc3O PAULO", ["sao paulo", "s\xe3o paulo", "tricolor paulista"]), addAliases_(e, "SANTOS", ["santos", "peixe", "taciano", "vila belmiro"]), addAliases_(e, "VASCO", ["vasco", "cruzmaltino", "cruz-maltino"]), addAliases_(e, "FLUMINENSE", ["fluminense", "flu", "tricolor carioca"]), addAliases_(e, "BOTAFOGO", ["botafogo", "glorioso", "artur cabral", "fogao", "fog\xe3o", "angelote"]), addAliases_(e, "INTERNACIONAL", ["internacional", "colorado", "beira rio", "beira-rio", "abel braga", "abelao", "abel\xe3o", "renata fan", "renata f\xe3", "alan patrick", "carboneiro", "loirao", "loir\xe3o", "edu dracena", "dracena", "capita"]), addAliases_(e, "GR\xcaMIO", ["gremio", "gr\xeamio", "tricolor gaucho", "tricolor ga\xfacho"]), addAliases_(e, "ATL\xc9TICO-MG", ["atletico mg", "atl\xe9tico mg", "galo", "atletico mineiro", "atl\xe9tico mineiro"]), addAliases_(e, "CRUZEIRO", ["cruzeiro", "raposa"]), addAliases_(e, "BAHIA", ["bahia", "tricolor de aco", "tricolor de a\xe7o"]), addAliases_(e, "VIT\xd3RIA", ["vitoria ba", "vit\xf3ria ba", "vitoria da bahia", "vit\xf3ria da bahia", "ec vitoria", "e.c. vitoria", "barradao", "barrad\xe3o", "leao da barra", "le\xe3o da barra", "toca do leao", "toca do le\xe3o", "jair ventura", "rubro negro baiano", "rubro-negro baiano"]), addAliases_(e, "FORTALEZA", ["fortaleza", "leao do pici", "le\xe3o do pici", "voivoda", "castelao", "castel\xe3o"]), addAliases_(e, "CEAR\xc1", ["ceara", "cear\xe1", "voz\xe3o", "vozao", "vovo", "vov\xf4"]), addAliases_(e, "ATHLETICO-PR", ["athletico", "athletico pr", "athletico-pr", "furacao", "furac\xe3o"]), addAliases_(e, "CORITIBA", ["coritiba", "coxabranca", "coxa"]), addAliases_(e, "GOI\xc1S", ["goias", "goi\xe1s"]), addAliases_(e, "SPORT", ["sport", "leao da ilha", "le\xe3o da ilha"]), addAliases_(e, "BRAGANTINO", ["bragantino", "red bull bragantino", "rb bragantino", "john john", "cleiton"]), addAliases_(e, "MILAN", ["milan", "rossonero"]), addAliases_(e, "TORINO", ["torino"]), addAliases_(e, "MANCHESTER UNITED", ["manchester united", "red devils", "overhampton"]), addAliases_(e, "LIVERPOOL", ["liverpool"]), addAliases_(e, "INTER DE MIL\xc3O", ["inter de milao", "inter de mil\xe3o", "neroazurra", "saniro"]), addAliases_(e, "F1", ["formula 1", "f1", "verstappen", "leclerc", "sainz", "norris", "piastri", "piastre", "bortoleto", "pit stop", "pits", "lando norris", "mclaren", "red bull racing", "campeao mundial de formula", "campe\xe3o mundial de f\xf3rmula"]), addAliases_(e, "NFL", ["monday night football", "super bowl"]), addAliases_(e, "OUTROS_ESPORTES", ["copa truck", "truck", "formula e", "f\xf3rmula e", "perdoncini", "podio", "p\xf3dio", "largada", "grid invertido", "relargada"]), addAliases_(e, "NBA", ["nba"]), addAliases_(e, "FUTSAL", ["futsal"]), addAliases_(e, "V\xd4LEI", ["volei", "v\xf4lei"]), addAliases_(e, "MMA", ["mma", "ufc"]),
    // --- Copa do Mundo 2026 / seleções / personagens (taxonomia Gazeta) ---
    addAliases_(e, "COPA DO MUNDO", ["copa do mundo", "copa 2026", "mundial", "world cup", "jogo da copa", "sede da copa", "estadio da copa", "est\xe1dio da copa", "giro copa", "giro da copa"]),
    addAliases_(e, "SELE\xc7\xc3O BRASILEIRA", ["selecao brasileira", "sele\xe7\xe3o brasileira", "selecao", "sele\xe7\xe3o", "amarelinha", "canarinho", "cbf", "convocacao", "convoca\xe7\xe3o", "camisa 10 do brasil"]),
    addAliases_(e, "ARGENTINA", ["argentina", "messi", "hermanos", "albiceleste", "scaloni"]),
    addAliases_(e, "FRAN\xc7A", ["franca", "fran\xe7a", "mbappe", "mbapp\xe9", "les bleus", "deschamps"]),
    addAliases_(e, "INGLATERRA", ["inglaterra", "english team", "kane", "bellingham", "tuchel"]),
    addAliases_(e, "NORUEGA", ["noruega", "haaland", "odegaard", "\xf8degaard"]),
    addAliases_(e, "M\xc9XICO", ["mexico", "m\xe9xico", "el tri", "azteca"]),
    addAliases_(e, "FIFA", ["fifa", "infantino", "fair play financeiro", "ranking fifa"]),
    addAliases_(e, "NEYMAR", ["neymar", "ney", "pai do neymar", "neymar pai"]),
    addAliases_(e, "ANCELOTTI", ["ancelotti", "carlo ancelotti", "carletto"]),
    addAliases_(e, "ARBITRAGEM", ["arbitragem", "var", "penalti nao marcado", "p\xeanalti n\xe3o marcado", "erro de arbitragem", "impedimento milimetrico", "impedimento milim\xe9trico"]),
    addAliases_(e, "FUTEBOL INTERNACIONAL", ["champions league", "liga dos campeoes", "liga dos campe\xf5es", "premier league", "la liga", "serie a italiana", "bundesliga", "real madrid", "barcelona", "psg", "manchester city", "bayern"]), {
        aliasToTheme: e
    }
}

function addAliases_(e, t, o) {
    for (const a of o) {
        const o = normalizeText_(a);
        o && (e[o] = t)
    }
}

// --- Tokenizers ---

function tokenize_(e, t) {
    const o = normalizeText_(e);
    if (!o) return [];
    const a = o.split(/[^a-z0-9]+/g),
        r = [];
    for (const e of a) e.length <= 2 || t.has(e) || /^\d+$/.test(e) || r.push(e);
    return r
}

function tokenizeWithAccents_(e, t) {
    if (!e) return {
        tokens: [],
        accentMap: new Map
    };
    const o = e.toString().toLowerCase(),
        a = normalizeText_(e),
        r = o.split(/[^a-z\xe1\xe0\xe2\xe3\xe9\xe8\xea\xed\xec\xee\xf3\xf2\xf4\xf5\xfa\xf9\xfb\xe70-9]+/gi).filter(e => e.length > 0),
        n = a.split(/[^a-z0-9]+/g).filter(e => e.length > 0),
        s = [],
        i = new Map;
    for (let e = 0; e < Math.min(r.length, n.length); e++) {
        const o = r[e],
            a = n[e];
        a.length <= 2 || t.has(a) || /^\d+$/.test(a) || (s.push(a), (!i.has(a) || o.length > i.get(a).length) && i.set(a, o))
    }
    return {
        tokens: s,
        accentMap: i
    }
}

// --- Desambiguação "Vitória" (EC Vitória vs. resultado de jogo) ---

const VITORIA_CLUBE_PATTERNS = {
        explicit: ["ec vitoria", "e.c. vitoria", "e c vitoria", "vitoria ba", "vitoria-ba", "vitoria da bahia", "esporte clube vitoria"],
        aliases: ["leao da barra", "le\xe3o da barra", "rubro negro baiano", "rubro-negro baiano", "barradao", "barrad\xe3o", "toca do leao", "toca do le\xe3o", "jair ventura"],
        clubContextRegex: [/\bo vitoria\s+(venceu|empatou|perdeu|ganhou|goleou|derrotou|enfrentou|jogou|enfrenta|joga|pega)/i, /\b(do|contra|pelo|para o|ao)\s+vitoria\b/i, /\bvitoria\s+(x|vs\.?|versus|contra)\s+/i, /\b(contra|x|vs\.?|versus)\s+vitoria\b/i, /\bjogador(es)?\s+(do|da)\s+vitoria\b/i, /\btorcida\s+(do|da)\s+vitoria\b/i, /\btorcedor(es)?\s+(do|da)\s+vitoria\b/i]
    },
    VITORIA_RESULTADO_PATTERNS = {
        verbalPatterns: [/\b(buscar|garantir|conquistar|conseguir|alcancar|alcan\xe7ar)\s+(a\s+)?vitoria\b/i, /\b(precisa|precisam|precisava|precisando)\s+(da|de uma|de)\s+vitoria\b/i, /\b(quer|querem|queria|querendo)\s+(a|uma)\s+vitoria\b/i, /\buma vitoria\s+(importante|crucial|fundamental|decisiva)\b/i, /\bvitoria\s+(por|de)\s+\d+\s*(a|x)\s*\d+/i, /\bvitoria\s+(em casa|fora|simples|magra|goleada)/i],
        pluralPatterns: [/\bvitorias\b/i, /\b(duas|tres|tr\xeas|quatro|cinco|\d+)\s+vitorias?\b/i],
        objectPatterns: [/\b(veio|vem|veio|chegou)\s+a\s+vitoria\b/i, /\ba vitoria\s+(veio|vem|chegou|aconteceu)\b/i]
    };

function classificarVitoria_(e, t, o = null) {
    const a = normalizeText_(e);
    if (!a.includes("vitoria")) return null;
    for (const e of VITORIA_CLUBE_PATTERNS.explicit)
        if (a.includes(normalizeText_(e))) return "EC_VITORIA";
    for (const e of VITORIA_CLUBE_PATTERNS.aliases)
        if (a.includes(normalizeText_(e))) return "EC_VITORIA";
    for (const e of VITORIA_CLUBE_PATTERNS.clubContextRegex)
        if (e.test(a)) return "EC_VITORIA";
    for (const e of VITORIA_RESULTADO_PATTERNS.verbalPatterns)
        if (e.test(a)) return "VITORIA_RESULTADO";
    for (const e of VITORIA_RESULTADO_PATTERNS.pluralPatterns)
        if (e.test(a)) return "VITORIA_RESULTADO";
    for (const e of VITORIA_RESULTADO_PATTERNS.objectPatterns)
        if (e.test(a)) return "VITORIA_RESULTADO";
    if ("VIT\xd3RIA" === t || "VITORIA" === t) return "EC_VITORIA";
    if (o) {
        const {
            temaPrev: e,
            temaProx: t
        } = o;
        if ("VIT\xd3RIA" === e || "VIT\xd3RIA" === t) return "EC_VITORIA"
    }
    return "VITORIA_RESULTADO"
}

function desambiguarTermosVitoria_(e, t, o) {
    return e.map(e => {
        const a = normalizeText_(e);
        return "vitoria" !== a && "vitorias" !== a ? e : "vitorias" === a ? "Vit\xf3ria (resultado)" : "EC_VITORIA" === classificarVitoria_(t, o) ? "EC Vit\xf3ria" : "Vit\xf3ria (resultado)"
    })
}

// --- Contagem de termos por minuto e top 20 global ---

function computeTopTermsPerMinute_(e, t) {
    const o = getEntityLexicon_(),
        a = [],
        r = new Map,
        n = new Map,
        s = new Map;
    for (let i = 0; i < e.length; i++) {
        const c = e[i].texto || "";
        if (e[i].isSuppressedMerchan) {
            a.push({
                terms: [],
                counts: new Map
            });
            continue
        }
        const l = detectEntityHits_(normalizeText_(c), o),
            {
                tokens: d,
                accentMap: u
            } = tokenizeWithAccents_(c, t),
            m = new Map;
        for (const e of d) m.set(e, (m.get(e) || 0) + 1);
        for (const [e, t] of u.entries()) s.has(e) || s.set(e, t);
        for (const e of l.hits) {
            const o = normalizeText_(e.alias);
            o && !t.has(o) && (m.set(o, (m.get(o) || 0) + 3), s.has(o) || s.set(o, e.alias))
        }
        const p = [...m.entries()].sort((e, t) => t[1] - e[1]).slice(0, CFG.TOP_TERMS_PER_MINUTE).map(e => e[0]);
        a.push({
            terms: p,
            counts: m
        });
        for (const [e, t] of m.entries()) r.set(e, (r.get(e) || 0) + t), n.has(e) || n.set(e, new Set), n.get(e).add(i)
    }
    return {
        minuteTop: a,
        globalCounts: r,
        globalMinutes: n,
        globalAccentMap: s
    }
}

function computeTop20_(e, t, o) {
    const {
        globalCounts: a,
        globalMinutes: r,
        globalAccentMap: n
    } = t, s = {
        count: 0,
        minutes: new Set
    }, i = {
        count: 0,
        minutes: new Set
    };
    for (let t = 0; t < e.length; t++) {
        const o = e[t].texto || "",
            a = normalizeText_(o);
        if (a.includes("vitoria") || a.includes("vitorias")) {
            const e = classificarVitoria_(o, null),
                r = (a.match(/vitoria/g) || []).length;
            "EC_VITORIA" === e ? (s.count += r, s.minutes.add(t)) : (i.count += r, i.minutes.add(t))
        }
    }
    const c = [];
    for (const [e, t] of a.entries()) {
        if (!e || o.has(e)) continue;
        const a = normalizeText_(e);
        if ("vitoria" === a || "vitorias" === a) continue;
        const s = r.get(e)?.size || 0,
            i = n.get(e) || e;
        c.push({
            term: i,
            termNorm: e,
            count: t,
            minutes: s
        })
    }
    return s.count > 0 && c.push({
        term: "EC Vit\xf3ria",
        termNorm: "ec_vitoria",
        count: s.count,
        minutes: s.minutes.size
    }), i.count > 0 && c.push({
        term: "Vit\xf3ria (resultado)",
        termNorm: "vitoria_resultado",
        count: i.count,
        minutes: i.minutes.size
    }), c.sort((e, t) => t.count - e.count || t.minutes - e.minutes || e.term.localeCompare(t.term)), c.slice(0, 20)
}