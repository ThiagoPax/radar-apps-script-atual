// =============================================================================
// 03_Timeline.js — Construção da timeline: MERCHAN, classificação, subtemas, suavização
// =============================================================================

// --- Construção principal da timeline (Jogo Aberto) ---

function buildTimeline_(e, t, o) {
    const a = getEntityLexicon_(),
        r = Array(e.length),
        n = [];
    let s = "GERAL",
        i = -999,
        c = 0,
        l = -1;
    for (let o = 0; o < e.length; o++) {
        const d = e[o].texto || "",
            u = t.minuteTop[o]?.terms || [];
        if (e[o].isSuppressedMerchan) {
            r[o] = {
                hora: e[o].horaStr,
                tema: "MERCHAN",
                subtema: "SUPRIMIDO",
                contentType: "MERCHAN",
                termos: "",
                trecho: "[Conteúdo comercial suprimido da transcrição]"
            };
            continue
        }
        const m = normalizeText_(d),
            p = detectMerchanAtStart_(m),
            g = detectThemeFromEntities_(m, a),
            A = detectCommercialSignals_(m);
        let E, S, h, f = !1,
            T = "";
        if (p ? (f = !0, T = "STARTS_WITH") : A.hasExplicitMerchan ? (f = !0, T = "EXPLICIT", A.startsMerchanBlock && (l = Math.max(l, o + CFG.MERCHAN_CARRY_MINUTES))) : A.totalScore >= CFG.MERCHAN_IMPLICIT_THRESHOLD ? detectAntiMerchanSignals_(m, g).hasStrongSportsContent || (f = !0, T = "HIGH_SCORE") : o <= l && A.totalScore >= 4 && (detectAntiMerchanSignals_(m, g).hasAnySportsContent || (f = !0, T = "CARRY")), f) E = "MERCHAN", S = merchanSubtema_(m, A), h = .95;
        else if (E = g.theme, S = ruleBasedSubtema_(m, g.secondaryTheme), h = g.confidence, "VIT\xd3RIA" === E && "EC_VITORIA" !== classificarVitoria_(d, null) && (E = "GERAL", h = .25, g.secondaryTheme && "VIT\xd3RIA" !== g.secondaryTheme && (E = g.secondaryTheme, h = .55)), S = tentarSubtemaPessoa_(m, u, S), "GERAL" === E) {
            const e = o - i;
            e >= 1 && e <= CFG.THEME_CARRY_MINUTES && c >= CFG.CARRY_MIN_CONFIDENCE && (E = s, h = Math.max(h, .55), S = "RESENHA")
        }
        "GERAL" !== E && "MERCHAN" !== E && h >= .6 && (s = E, i = o, c = h), !f && ("GERAL" === E || h < .5) && CFG.LLM_ENABLED && n.push({
            idx: o,
            hora: e[o].horaStr,
            text: d,
            topTerms: u
        }), r[o] = {
            hora: e[o].horaStr,
            tema: E,
            subtema: S,
            contentType: E,
            termos: u.join(", "),
            trecho: d,
            _conf: h,
            _merchanScore: A.totalScore,
            _merchanReason: T
        }
    }
    if (CFG.LLM_ENABLED && n.length > 0 && hasOpenAIKey_()) {
        const e = classifyWithLLM_(n);
        for (let t = 0; t < r.length; t++) {
            if ("MERCHAN" === r[t].tema) continue;
            const o = t + "";
            if (!e.has(o)) continue;
            const a = e.get(o),
                n = sanitizeTheme_(a.tema);
            n && (r[t].tema = n);
            const s = sanitizeSubtema_(a.subtema);
            s && (r[t].subtema = s)
        }
    }
    applyMerchanSandwichRule_(r), smoothTimelineThemes_(r);
    for (const e of r) delete e._conf, delete e._merchanScore, delete e._merchanReason;
    return r
}

// --- Detecção de MERCHAN ---

function detectMerchanAtStart_(e) {
    const t = e.substring(0, 100),
        o = ["eu to falando dele", "eu t\xf4 falando dele", "magnesio", "magn\xe9sio", "francis life", "bajage", "pulsar", "estival", "band shop", "bandshop", "voce sente que a sua energia", "que tal economizar"];
    for (const e of o)
        if (t.includes(normalizeText_(e))) return !0;
    return !1
}

function applyMerchanSandwichRule_(e) {
    const t = e.length;
    for (let o = 1; o < t - 1; o++) {
        if ("MERCHAN" === e[o].tema) continue;
        if ("SUPRIMIDO" === e[o].subtema) continue;
        const t = e[o - 1].tema,
            a = e[o + 1].tema;
        if ("MERCHAN" === t && "MERCHAN" === a) {
            const t = normalizeText_(e[o].trecho || ""),
                a = detectCommercialSignals_(t);
            e[o].tema = "MERCHAN", e[o].subtema = merchanSubtema_(t, a), e[o].contentType = "MERCHAN"
        }
    }
    for (let o = 1; o < t - 2; o++) {
        if ("MERCHAN" === e[o].tema) continue;
        if ("MERCHAN" === e[o + 1].tema) continue;
        const t = e[o - 1].tema,
            a = e[o + 2].tema;
        if ("MERCHAN" === t && "MERCHAN" === a) {
            const t = normalizeText_(e[o].trecho || ""),
                a = normalizeText_(e[o + 1].trecho || ""),
                r = detectCommercialSignals_(t),
                n = detectCommercialSignals_(a);
            (r.totalScore >= 4 || n.totalScore >= 4 || r.hasExplicitMerchan || n.hasExplicitMerchan) && (e[o].tema = "MERCHAN", e[o].subtema = merchanSubtema_(t, r), e[o].contentType = "MERCHAN", e[o + 1].tema = "MERCHAN", e[o + 1].subtema = merchanSubtema_(a, n), e[o + 1].contentType = "MERCHAN")
        }
    }
}

// --- Detecção de tema via entidades ---

function detectThemeFromEntities_(e, t) {
    const o = detectEntityHits_(e, t);
    let a = o.bestTheme || "GERAL",
        r = o.bestScore || 0,
        n = .25;
    "GERAL" !== a && (n = Math.min(.95, .5 + r / 8));
    const s = o.secondaryTheme || "";
    if ("GERAL" === a) {
        if (containsAny_(e, ["formula 1", "f1", "verstappen", "norris", "piastri", "mclaren", "pole position"])) return {
            theme: "F1",
            confidence: .75,
            secondaryTheme: ""
        };
        if (containsAny_(e, ["monday night football", "super bowl", "touchdown"])) return {
            theme: "NFL",
            confidence: .7,
            secondaryTheme: ""
        };
        if (containsAny_(e, ["copa truck", "truck", "perdoncini"])) return {
            theme: "OUTROS_ESPORTES",
            confidence: .75,
            secondaryTheme: ""
        };
        if (containsAny_(e, ["champions league", "champions"])) {
            if (containsAny_(e, ["liverpool"])) return {
                theme: "LIVERPOOL",
                confidence: .7,
                secondaryTheme: ""
            };
            if (containsAny_(e, ["inter de milao", "inter de mil\xe3o", "saniro"])) return {
                theme: "INTER DE MIL\xc3O",
                confidence: .7,
                secondaryTheme: ""
            }
        }
    }
    return {
        theme: a,
        confidence: n,
        secondaryTheme: s
    }
}

function detectEntityHits_(e, t) {
    const o = t.aliasToTheme,
        a = new Map,
        r = [],
        n = new Map;
    for (const t in o) {
        if (!t) continue;
        const s = t.length <= 5;
        let i = -1;
        if (s) {
            const o = RegExp(`\\b${escapeRegex_(t)}\\b`, "i");
            o.test(e) && (i = e.search(o))
        } else i = e.indexOf(t);
        if (i >= 0) {
            const e = o[t],
                c = s ? 2 : 3;
            a.set(e, (a.get(e) || 0) + c), r.push({
                alias: t,
                theme: e,
                w: c
            }), (!n.has(e) || i < n.get(e)) && n.set(e, i)
        }
    }
    let s = "",
        i = 0;
    for (const [e, t] of a.entries()) t > i ? (i = t, s = e) : t === i && t > 0 && (n.get(e) ?? 999999) < (n.get(s) ?? 999999) && (s = e);
    let c = "",
        l = 0;
    for (const [e, t] of a.entries()) e !== s && t > l && (l = t, c = e);
    return {
        bestTheme: s,
        bestScore: i,
        secondaryTheme: c,
        hits: r
    }
}

// --- Detecção de sinais comerciais / anti-MERCHAN ---

function detectCommercialSignals_(e) {
    let t = 0,
        o = 0,
        a = !1;
    return containsAny_(e, ["aponte sua camera", "aponte a camera", "aponte sua c\xe2mera", "aponte a c\xe2mera", "qr code", "qrcode"]) && (t += 10, a = !0), containsAny_(e, ["com oferecimento", "oferecimento bandbet", "oferecimento band bet", "band placar com oferecimento"]) && (t += 10, a = !0), (/r\$\s*\d+[\.,]?\d*/.test(e) || containsAny_(e, ["por apenas", "parcelas de"])) && (t += 8, a = !0), (/0800\s*\d{3,}/.test(e) || containsAny_(e, ["ligue agora", "ligue para"])) && (t += 10, a = !0), containsAny_(e, ["bandbet.bet", "budbet.bet", ".com.br", "bandshop.com"]) && (t += 8, a = !0), containsAny_(e, ["use o cupom", "cupom band", "cupom"]) && (t += 6, a = !0), containsAny_(e, ["jogar com responsabilidade", "maiores de 18"]) && (t += 8, a = !0), containsAny_(e, ["bajage", "pulsar n150", "estival", "francis life", "band shop", "bandshop", "lavadora de alta pressao", "magnesio", "magn\xe9sio"]) && (o += 8), containsAny_(e, ["melhores oportunidades do mercado", "odds", "cotacao", "cota\xe7\xe3o", "budbet", "bandbet"]) && (o += 5), containsAny_(e, ["acesse", "confira", "garanta", "aproveite", "desconto", "frete gratis", "frete gr\xe1tis"]) && (o += 3), {
        explicitScore: t,
        implicitScore: o,
        totalScore: t + o,
        hasExplicitMerchan: a,
        startsMerchanBlock: containsAny_(e, ["band placar com oferecimento", "oferecimento bandbet", "oferecimento band bet"]) || t >= 15
    }
}

function detectAntiMerchanSignals_(e, t) {
    let o = 0,
        a = !1,
        r = !1;
    "GERAL" !== t.theme && "MERCHAN" !== t.theme && (o += 8, r = !0, t.confidence >= .55 && (a = !0), t.secondaryTheme && "GERAL" !== t.secondaryTheme && "MERCHAN" !== t.secondaryTheme && (o += 10, a = !0));
    const n = ["flamengo", "palmeiras", "corinthians", "sao paulo", "s\xe3o paulo", "santos", "vasco", "fluminense", "botafogo", "internacional", "gremio", "gr\xeamio", "atletico", "atl\xe9tico", "cruzeiro", "bahia", "fortaleza", "ceara", "cear\xe1", "bragantino", "athletico", "coritiba", "goias", "goi\xe1s", "sport", "manchester", "liverpool", "barcelona", "real madrid", "milan", "inter de milao", "juventus", "bayern", "psg", "chelsea", "arsenal", "torino", "napoli"].filter(t => e.includes(t));
    n.length >= 1 && (o += 8, r = !0), n.length >= 2 && (o += 8, a = !0);
    const s = ["gol", "placar", "venceu", "derrotou", "empat", "virou", "rebaixamento", "serie a", "s\xe9rie a", "serie b", "s\xe9rie b", "libertadores", "sul-americana", "copa do brasil", "primeiro tempo", "segundo tempo", "apito final", "campeonato", "rodada", "pontos", "tabela", "classificacao", "classifica\xe7\xe3o", "tecnico", "t\xe9cnico", "treinador", "atacante", "zagueiro", "meia", "volante", "goleiro", "premier league", "champions league", "la liga", "bundesliga", "italiano", "penalti", "p\xeanalti", "virada", "goleada", "balancou a rede", "balan\xe7ou a rede", "titulo", "t\xedtulo", "campeao", "campe\xe3o", "vice", "selecao", "sele\xe7\xe3o", "premiacao", "premia\xe7\xe3o", "artilheiro", "artilharia", "escalado", "convocado", "lesao", "les\xe3o", "contusao", "contus\xe3o"];
    return containsAny_(e, s) && (o += 6, r = !0, countMatches_(e, s) >= 2 && (a = !0)), containsAny_(e, ["vai cair", "vai descer", "nao cai", "n\xe3o cai", "permanencia", "perman\xeancia", "escapou"]) && (o += 4, r = !0), containsAny_(e, ["abel braga", "abelao", "abel\xe3o", "abel ferreira", "neymar", "alan patrick", "carboneiro", "jair ventura", "voivoda", "mano menezes", "arrascaeta", "gabigol", "dudu", "flaco lopez", "veiga", "bruno fernandes", "bellard", "rabiot", "zapata", "mancini", "cuca", "renata fan", "denilson"]) && (o += 6, r = !0), {
        score: o,
        hasStrongSportsContent: a,
        hasAnySportsContent: r
    }
}

function countMatches_(e, t) {
    let o = 0;
    for (const a of t) e.includes(normalizeText_(a)) && o++;
    return o
}

// --- Subtemas ---

function merchanSubtema_(e, t) {
    return containsAny_(e, ["voltamos ja", "voltamos j\xe1", "intervalo comercial"]) ? "BREAK" : containsAny_(e, ["bandbet", "band bet", "budbet", "odds", "cotacao", "cota\xe7\xe3o", "melhores oportunidades", "jogar com responsabilidade"]) ? "BAND_BET" : containsAny_(e, ["bajage", "pulsar", "estival", "francis life", "magnesio", "magn\xe9sio", "vitamina", "suplemento"]) || containsAny_(e, ["band shop", "bandshop"]) ? "PRODUTO" : containsAny_(e, ["oferecimento", "apresentado por"]) ? "PATROCINIO" : "CTA"
}

function ruleBasedSubtema_(e, t) {
    return t && "MERCHAN" !== t && "GERAL" !== t && containsAny_(e, ["contra", "versus", "x ", " x ", "enfrenta", "pega o", "pega a"]) ? t : containsAny_(e, ["penalti", "p\xeanalti", "arbitro", "\xe1rbitro", "arbitragem", "var"]) ? "ARBITRAGEM" : containsAny_(e, ["rebaix", "serie b", "s\xe9rie b", "cair", "caiu", "escapou", "z4", "vai descer"]) ? "REBAIXAMENTO" : containsAny_(e, ["classific", "lider", "l\xedder", "vaga", "g4", "g-4", "libertadores", "sul-americana", "champions"]) ? "CLASSIFICACAO" : containsAny_(e, ["gol", "placar", "venceu", "derrotou", "empate", "virou", "vitoria", "vit\xf3ria", "apito final"]) ? "JOGO" : containsAny_(e, ["bastidor", "resenha", "zoeira", "risadas"]) ? "BASTIDOR" : containsAny_(e, ["torcida", "torcedor"]) ? "TORCIDA" : containsAny_(e, ["palpite", "quem cai"]) ? "PALPITE" : containsAny_(e, ["formula 1", "f1", "verstappen", "norris", "pole", "grid", "corrida", "largada"]) || containsAny_(e, ["copa truck", "truck"]) ? "CORRIDA" : "OUTROS"
}

function tentarSubtemaPessoa_(e, t, o) {
    if ("OUTROS" !== o) return o;
    const a = [];
    for (const o of t) {
        const t = normalizeText_(o);
        PESSOAS_FUTEBOL[t] && a.push({
            nome: PESSOAS_FUTEBOL[t],
            termo: t,
            fonte: "base"
        });
        const r = detectarNomeDinamico_(t, e);
        r && a.push({
            nome: r,
            termo: t,
            fonte: "dinamico"
        })
    }
    if (0 === a.length) return o;
    for (const t of a) {
        const o = contarOcorrencias_(e, t.termo),
            a = verificarContextoPessoa_(e, t.termo);
        if (o >= 2 || a) return t.nome
    }
    return o
}

const PESSOAS_FUTEBOL = {
    "abel ferreira": "Abel Ferreira",
    abel: "Abel Ferreira",
    "roger machado": "Roger Machado",
    "fernando diniz": "Fernando Diniz",
    diniz: "Fernando Diniz",
    "mano menezes": "Mano Menezes",
    mano: "Mano Menezes",
    tite: "Tite",
    "dorival junior": "Dorival Júnior",
    dorival: "Dorival Júnior",
    "filipe luis": "Filipe Luís",
    cuca: "Cuca",
    vojvoda: "Vojvoda",
    "ramon diaz": "Ramón Díaz",
    "jair ventura": "Jair Ventura",
    "artur jorge": "Artur Jorge",
    "fernando seabra": "Fernando Seabra",
    "renato gaucho": "Renato Gaúcho",
    "pedro caixinha": "Pedro Caixinha",
    "rogerio ceni": "Rogério Ceni",
    neymar: "Neymar",
    gabigol: "Gabigol",
    "gabriel barbosa": "Gabigol",
    arrascaeta: "Arrascaeta",
    dudu: "Dudu",
    endrick: "Endrick",
    "vini jr": "Vini Jr",
    "vinicius junior": "Vini Jr",
    raphinha: "Raphinha",
    richarlison: "Richarlison",
    casemiro: "Casemiro",
    "alan patrick": "Alan Patrick",
    borre: "Borré",
    "bruno henrique": "Bruno Henrique",
    "everton ribeiro": "Éverton Ribeiro",
    "raphael veiga": "Raphael Veiga",
    veiga: "Raphael Veiga",
    "flaco lopez": "Flaco López",
    rony: "Rony",
    "yuri alberto": "Yuri Alberto",
    romero: "Romero",
    "hugo souza": "Hugo Souza",
    "memphis depay": "Memphis Depay",
    memphis: "Memphis Depay",
    calleri: "Calleri",
    "lucas moura": "Lucas Moura",
    oscar: "Oscar",
    "luiz henrique": "Luiz Henrique",
    savarino: "Savarino",
    "thiago almada": "Thiago Almada",
    "igor jesus": "Igor Jesus",
    "matheus pereira": "Matheus Pereira",
    "renata fan": "Renata Fan",
    denilson: "Denilson",
    "chico garcia": "Chico Garcia",
    "ronaldo giovanelli": "Ronaldo Giovanelli",
    velloso: "Velloso",
    "ulisses costa": "Ulisses Costa",
    "fabio sormani": "Fábio Sormani",
    sormani: "Sormani",
    "leila pereira": "Leila Pereira",
    "rodolfo landim": "Rodolfo Landim",
    "john textor": "John Textor",
    "augusto melo": "Augusto Melo",
    daronco: "Anderson Daronco",
    "anderson daronco": "Anderson Daronco",
    "rafael claus": "Rafael Claus"
};

function detectarNomeDinamico_(e, t) {
    if (e.length < 4) return null;
    const o = ["tecnico " + e, "t\xe9cnico " + e, "jogador " + e, e + " disse", e + " falou", e + " marcou", e + " fez", "gol de " + e, "gol do " + e, e + " entrou", e + " saiu"];
    for (const a of o)
        if (t.includes(a)) return capitalizarNome_(e);
    return null
}

function capitalizarNome_(e) {
    return e.split(" ").map(e => e.charAt(0).toUpperCase() + e.slice(1)).join(" ")
}

function contarOcorrencias_(e, t) {
    const o = RegExp(t, "gi"),
        a = e.match(o);
    return a ? a.length : 0
}

function verificarContextoPessoa_(e, t) {
    const o = ["tecnico " + t, "t\xe9cnico " + t, "treinador " + t, "jogador " + t, "atacante " + t, "meia " + t, "volante " + t, "zagueiro " + t, "lateral " + t, "goleiro " + t, t + " disse", t + " falou", t + " declarou", t + " afirmou", "segundo " + t, "para " + t, "de " + t, "do " + t, "gol de " + t, "gol do " + t];
    for (const t of o)
        if (e.includes(t)) return !0;
    return !1
}

// --- Sanitização de tema/subtema vindo do LLM ---

function sanitizeTheme_(e) {
    const t = (e || "").toString().trim();
    if (!t) return null;
    const o = stripAccents_(t).toUpperCase();
    for (const e of CFG.FIXED_THEMES)
        if (stripAccents_(e).toUpperCase() === o) return e;
    return {
        "SAO PAULO": "SÃO PAULO",
        "ATLETICO MG": "ATLÉTICO-MG",
        "ATLETICO-MG": "ATLÉTICO-MG",
        GREMIO: "GRÊMIO",
        CEARA: "CEARÁ",
        VITORIA: "VITÓRIA",
        GOIAS: "GOIÁS",
        INTER: "INTERNACIONAL"
    } [o] || null
}

function sanitizeSubtema_(e) {
    const t = (e || "").toString().trim();
    if (!t) return null;
    let o = t.replace(/\s+/g, " ").trim();
    return o.length > 24 && (o = o.slice(0, 24).trim()), o.toUpperCase()
}

// --- Suavização da timeline ---

function smoothTimelineThemes_(e) {
    const t = e.length,
        o = CFG.SMOOTH_WINDOW;

    function a(a, r = !0) {
        const n = new Map;
        for (let s = Math.max(0, a - o); s <= Math.min(t - 1, a + o); s++) {
            const t = e[s].tema;
            t && "GERAL" !== t && (r && "MERCHAN" === t || n.set(t, (n.get(t) || 0) + 1))
        }
        let s = "",
            i = 0;
        for (const [e, t] of n.entries()) t > i && (i = t, s = e);
        return {
            theme: s,
            count: i
        }
    }
    for (let o = 1; o < t - 1; o++) {
        if ("GERAL" !== e[o].tema) continue;
        const t = e[o - 1].tema,
            a = e[o + 1].tema;
        t && a && t === a && "MERCHAN" !== t && (e[o].tema = t, e[o].subtema = "RESENHA")
    }
    for (let o = 1; o < t - 1; o++) {
        if ("MERCHAN" === e[o].tema) continue;
        if ("GERAL" === e[o].tema) continue;
        const t = e[o].tema,
            r = e[o - 1].tema,
            n = e[o + 1].tema;
        if (t !== r && t !== n && r === n && "GERAL" !== r) {
            const r = a(o);
            r.theme && r.theme !== t && r.count >= 2 && (e[o].tema = r.theme, e[o].subtema = "RESENHA")
        }
    }
    for (let o = 0; o < t; o++) {
        if ("GERAL" !== e[o].tema) continue;
        const t = a(o);
        t.theme && t.count >= 2 && (e[o].tema = t.theme, e[o].subtema = "RESENHA")
    }
    for (let o = 0; o < t; o++) e[o].contentType = e[o].tema
}