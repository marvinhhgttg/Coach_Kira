// ── Formel-Rechner OHNE eval/Function ──────────────────────────────
// Eigenes Tokenizer + AST-Parser, elementweise Auswertung über Tage.

// ── Token-Typen ───────────────────────────────────────────────────
type Token = { type: 'NUMBER' | 'IDENT' | 'OP' | 'LPAREN' | 'RPAREN' | 'COMMA' | 'EOF'; value?: number; name?: string; symbol?: string };

export type Knoten =
  | { typ: 'Number'; value: number }
  | { typ: 'Variable'; name: string }
  | { typ: 'BinOp'; operator: string; left: Knoten; right: Knoten }
  | { typ: 'UnaryMinus'; operand: Knoten }
  | { typ: 'FuncCall'; name: string; args: Knoten[] };

// ── Tokenizer ─────────────────────────────────────────────────────
export function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  while (i < src.length) {
    const ch = src[i];

    // Leerzeichen überspringen
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }

    // Zahlen (nur Punkt als Dezimaltrenner, keine Kommas in Zahlen)
    if (ch >= '0' && ch <= '9') {
      let numStr = '';
      while (i < src.length && ((src[i] >= '0' && src[i] <= '9') || src[i] === '.')) {
        numStr += src[i];
        i++;
      }
      const val = parseFloat(numStr);
      if (isNaN(val)) throw new Error(`Kann Zahl nicht parsen: "${numStr}"`);
      tokens.push({ type: 'NUMBER', value: val });
      continue;
    }

    // Punkt als Zahl? (allein)
    if (ch === '.') {
      // Nur wenn nicht gefolgt von Ziffer → könnte es ein Fehler sein
      if (i + 1 < src.length && src[i + 1] >= '0' && src[i + 1] <= '9') {
        let numStr = '';
        i++; // Punkt überspringen
        while (i < src.length && ((src[i] >= '0' && src[i] <= '9') || src[i] === '.')) {
          numStr += src[i];
          i++;
        }
        const val = parseFloat(numStr);
        if (isNaN(val)) throw new Error(`Kann Zahl nicht parsen: "${numStr}"`);
        tokens.push({ type: 'NUMBER', value: val });
      } else {
        throw new Error('Unerwartetes Zeichen am Punkt: "."');
      }
      continue;
    }

    // Bezeichner [A-Za-z_][A-Za-z0-9_]*
    if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || ch === '_') {
      let ident = '';
      while (i < src.length && ((src[i] >= 'a' && src[i] <= 'z') || (src[i] >= 'A' && src[i] <= 'Z') || (src[i] >= '0' && src[i] <= '9') || src[i] === '_')) {
        ident += src[i];
        i++;
      }
      tokens.push({ type: 'IDENT', name: ident });
      continue;
    }

    // Operatoren und Klammern
    if (ch === '+') { tokens.push({ type: 'OP', symbol: '+' }); i++; continue; }
    if (ch === '-') { tokens.push({ type: 'OP', symbol: '-' }); i++; continue; }
    if (ch === '*') { tokens.push({ type: 'OP', symbol: '*' }); i++; continue; }
    if (ch === '/') { tokens.push({ type: 'OP', symbol: '/' }); i++; continue; }
    if (ch === '^') { tokens.push({ type: 'OP', symbol: '^' }); i++; continue; }
    if (ch === '(') { tokens.push({ type: 'LPAREN' }); i++; continue; }
    if (ch === ')') { tokens.push({ type: 'RPAREN' }); i++; continue; }
    if (ch === ',') { tokens.push({ type: 'COMMA' }); i++; continue; }

    throw new Error(`Unbekanntes Zeichen "${ch}" an Position ${i}`);
  }

  tokens.push({ type: 'EOF' });
  return tokens;
}

// ── Parser (Pratt-Parsen für Operatorprezedenz) ────────────────────
class Parser {
  tokens: Token[];
  pos = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  current(): Token {
    return this.tokens[this.pos];
  }

  consume(type: Token['type']): Token {
    const tok = this.current();
    if (tok.type !== type) {
      throw new Error(`Erwartet Token vom Typ "${type}", habe "${tok.type}" an Position ${this.pos}`);
    }
    this.pos++;
    return tok;
  }

  // Parse eine vollständige Formel
  parse(): Knoten {
    const node = this.parseExpression();
    if (this.current().type !== 'EOF') {
      throw new Error(`Übriger Text nach Formel: "${this.tokens[this.pos].symbol ?? this.current().type}"`);
    }
    return node;
  }

  // Expressions: +, - (niedrigste Priorität)
  parseExpression(): Knoten {
    let left = this.parseTerm();
    while (this.current().type === 'OP' && (this.current().symbol === '+' || this.current().symbol === '-')) {
      const op = this.consume('OP').symbol;
      const right = this.parseTerm();
      left = { typ: 'BinOp', operator: op ?? '', left, right };
    }
    return left;
  }

  // Terms: *, / (höhere Priorität)
  parseTerm(): Knoten {
    let left = this.parseExponent();
    while (this.current().type === 'OP' && (this.current().symbol === '*' || this.current().symbol === '/')) {
      const op = this.consume('OP').symbol;
      const right = this.parseExponent();
      left = { typ: 'BinOp', operator: op ?? '', left, right };
    }
    return left;
  }

  // Exponent: ^ (höchste Priorität, rechtsassoziativ)
  parseExponent(): Knoten {
    const base = this.parseUnary();
    if (this.current().type === 'OP' && this.current().symbol === '^') {
      this.consume('OP'); // ^ konsumieren
      const right = this.parseExponent(); // rechtsassoziativ: rekursiv
      return { typ: 'BinOp', operator: '^', left: base, right };
    }
    return base;
  }

  // Unary: -x (unäres Minus)
  parseUnary(): Knoten {
    if (this.current().type === 'OP' && this.current().symbol === '-') {
      this.consume('OP');
      const operand = this.parseUnary();
      return { typ: 'UnaryMinus', operand };
    }
    // Auch unäres + erlaubt (überspringen)
    if (this.current().type === 'OP' && this.current().symbol === '+') {
      this.consume('OP');
      return this.parseUnary();
    }
    return this.parsePrimary();
  }

  // Primary: Number, Ident, (Expr), FunctionCall
  parsePrimary(): Knoten {
    const tok = this.current();

    if (tok.type === 'NUMBER') {
      this.consume('NUMBER');
      return { typ: 'Number', value: tok.value ?? 0 };
    }

    if (tok.type === 'IDENT') {
      this.consume('IDENT');
      // Nächster Token: LPAREN? → Funktionsaufruf
      if (this.current().type === 'LPAREN') {
        return this.parseFunctionCall(tok.name ?? '');
      }
      return { typ: 'Variable', name: tok.name ?? '' };
    }

    if (tok.type === 'LPAREN') {
      this.consume('LPAREN');
      const node = this.parseExpression();
      this.consume('RPAREN');
      return node;
    }

    throw new Error(`Unerwartetes Token "${tok.type}" an Position ${this.pos}`);
  }

  // Funktionsaufruf: name(arg1, arg2, ...)
  parseFunctionCall(name: string): Knoten {
    this.consume('LPAREN'); // ( konsumieren
    const args: Knoten[] = [];

    if (this.current().type !== 'RPAREN') {
      args.push(this.parseExpression());

      while (this.current().type === 'COMMA') {
        this.consume('COMMA'); // , konsumieren
        args.push(this.parseExpression());
      }
    }

    this.consume('RPAREN'); // ) konsumieren
    return { typ: 'FuncCall', name, args };
  }
}

// ── parseFormel: src → AST-Knoten ──────────────────────────────────
export function parseFormel(src: string): Knoten {
  const tokens = tokenize(src);
  const parser = new Parser(tokens);
  return parser.parse();
}

// ── rechne: AST + Reihennamen → Array<number|null> elementweise ────

type Reihen = Record<string, Array<number | null>>;

const FENSTER = ['avg', 'sum', 'min', 'max'];
const EINFACH = ['abs', 'sqrt', 'log'];

/** Prüft Variablen und Funktionen vorab, damit Tippfehler eine klare Meldung geben. */
function pruefe(k: Knoten, reihen: Reihen): void {
  switch (k.typ) {
    case 'Number':
      return;
    case 'Variable':
      if (!(k.name in reihen)) throw new Error(`Unbekannte Variable „${k.name}“. Erlaubt: ${Object.keys(reihen).join(', ')}`);
      return;
    case 'UnaryMinus':
      return pruefe(k.operand, reihen);
    case 'BinOp':
      pruefe(k.left, reihen);
      return pruefe(k.right, reihen);
    case 'FuncCall': {
      if (FENSTER.includes(k.name) || k.name === 'lag') {
        if (k.args.length !== 2) throw new Error(`${k.name}(x, n) braucht zwei Argumente`);
        if (k.args[1].typ !== 'Number' || k.args[1].value < (k.name === 'lag' ? 0 : 1)) throw new Error(`${k.name}: n muss eine feste Zahl sein`);
      } else if (EINFACH.includes(k.name)) {
        if (k.args.length !== 1) throw new Error(`${k.name}(x) braucht ein Argument`);
      } else {
        throw new Error(`Unbekannte Funktion „${k.name}“. Erlaubt: avg, sum, min, max, lag, abs, sqrt, log`);
      }
      k.args.forEach((a) => pruefe(a, reihen));
      return;
    }
  }
}

function wert(k: Knoten, reihen: Reihen, i: number): number | null {
  if (i < 0) return null;
  switch (k.typ) {
    case 'Number':
      return k.value;
    case 'Variable': {
      const v = reihen[k.name]?.[i];
      return v == null || !Number.isFinite(v) ? null : v;
    }
    case 'UnaryMinus': {
      const v = wert(k.operand, reihen, i);
      return v == null ? null : -v;
    }
    case 'BinOp': {
      const a = wert(k.left, reihen, i);
      const b = wert(k.right, reihen, i);
      if (a == null || b == null) return null;
      switch (k.operator) {
        case '+': return a + b;
        case '-': return a - b;
        case '*': return a * b;
        case '/': return b === 0 ? null : a / b;
        case '^': {
          const r = Math.pow(a, b);
          return Number.isFinite(r) ? r : null;
        }
        default: return null;
      }
    }
    case 'FuncCall': {
      const n = k.args[1] && k.args[1].typ === 'Number' ? Math.round(k.args[1].value) : 0;
      if (k.name === 'lag') return wert(k.args[0], reihen, i - n);
      if (FENSTER.includes(k.name)) {
        const xs: number[] = [];
        for (let j = Math.max(0, i - n + 1); j <= i; j++) {
          const v = wert(k.args[0], reihen, j);
          if (v != null) xs.push(v);
        }
        if (!xs.length) return null;
        if (k.name === 'avg') return xs.reduce((s, x) => s + x, 0) / xs.length;
        if (k.name === 'sum') return xs.reduce((s, x) => s + x, 0);
        if (k.name === 'min') return Math.min(...xs);
        return Math.max(...xs);
      }
      const v = wert(k.args[0], reihen, i);
      if (v == null) return null;
      if (k.name === 'abs') return Math.abs(v);
      if (k.name === 'sqrt') return v < 0 ? null : Math.sqrt(v);
      if (k.name === 'log') return v <= 0 ? null : Math.log(v);
      return null;
    }
  }
  return null;
}

/** Elementweise über alle Tage; Fensterfunktionen schauen nur zurück (inkl. heute). */
export function rechne(ast: Knoten, reihen: Reihen): Array<number | null> {
  pruefe(ast, reihen);
  const n = Math.max(0, ...Object.values(reihen).map((r) => r.length));
  const out: Array<number | null> = [];
  for (let i = 0; i < n; i++) {
    const v = wert(ast, reihen, i);
    out.push(v == null || !Number.isFinite(v) ? null : v);
  }
  return out;
}

export const VARIABLEN: Record<string, string> = {
  load: 'Tageslast (ESS)',
  atl: 'ATL – akute Belastung (7 Tage)',
  ctl: 'CTL – chronische Belastung, Fitness (28 Tage)',
  acwr: 'ACWR – ATL / CTL',
  rhr: 'Ruhepuls (bpm)',
  hrv: 'HRV, 7-Tage-Schnitt (ms)',
  sleep: 'Schlaf (Stunden)',
  sleepScore: 'Schlafscore (0–100)',
  readiness: 'Trainingsbereitschaft (0–100)',
  es: 'Ausdauerwert (Garmin Endurance Score)',
  mono: 'Monotonie (7 Tage)',
};
