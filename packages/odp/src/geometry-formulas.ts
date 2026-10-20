import type { CustomGeometryOptions } from "@office-open/core";

type FormulaParts = readonly string[];

export class GeometryFormulaEvaluator {
  private readonly results = new Map<string, number>();
  private readonly evaluating = new Set<string>();

  constructor(
    private readonly geometry: CustomGeometryOptions,
    private readonly width: number,
    private readonly height: number,
  ) {}

  public coordinate(value: string): number {
    const result = this.token(value);
    if (result === undefined || !Number.isFinite(result)) {
      throw new RangeError(`Unsupported geometry formula: ${value}`);
    }
    return result;
  }

  private guide(name: string): number | undefined {
    const cached = this.results.get(name);
    if (cached !== undefined) return cached;
    if (this.evaluating.has(name)) return undefined;
    const guide =
      this.geometry.guides?.find((candidate) => candidate.name === name) ??
      this.geometry.adjustmentValues?.find((candidate) => candidate.name === name);
    if (!guide) return undefined;
    this.evaluating.add(name);
    const value = this.formula(guide.formula.split(/\s+/u).filter(Boolean));
    this.evaluating.delete(name);
    if (value === undefined || !Number.isFinite(value)) return undefined;
    this.results.set(name, value);
    return value;
  }

  private token(value: string): number | undefined {
    if (/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?$/u.test(value)) return Number(value);
    if (value === "w") return this.width;
    if (value === "h") return this.height;
    if (value === "l" || value === "t") return 0;
    if (value === "r" || value === "b") return value === "r" ? this.width : this.height;
    if (value === "hc") return this.width / 2;
    if (value === "vc") return this.height / 2;
    if (value === "hd2") return this.height / 2;
    if (value === "hd3") return this.height / 3;
    if (value === "hd4") return this.height / 4;
    if (value === "hd5") return this.height / 5;
    if (value === "hd6") return this.height / 6;
    if (value === "hd8") return this.height / 8;
    if (value === "ls") return Math.max(this.width, this.height);
    if (value === "ss") return Math.min(this.width, this.height);
    if (value === "ssd2") return Math.min(this.width, this.height) / 2;
    if (value === "ssd4") return Math.min(this.width, this.height) / 4;
    if (value === "ssd6") return Math.min(this.width, this.height) / 6;
    if (value === "ssd8") return Math.min(this.width, this.height) / 8;
    if (value === "ssd16") return Math.min(this.width, this.height) / 16;
    if (value === "ssd32") return Math.min(this.width, this.height) / 32;
    if (value === "wd2") return this.width / 2;
    if (value === "wd3") return this.width / 3;
    if (value === "wd4") return this.width / 4;
    if (value === "wd5") return this.width / 5;
    if (value === "wd6") return this.width / 6;
    if (value === "wd8") return this.width / 8;
    if (value === "wd10") return this.width / 10;
    if (value === "wd32") return this.width / 32;
    if (value === "cd2") return 5_400_000;
    if (value === "cd4") return 2_700_000;
    if (value === "cd8") return 1_350_000;
    if (value === "3cd4") return 16_200_000;
    if (value === "3cd8") return 4_050_000;
    if (value === "5cd8") return 6_750_000;
    if (value === "7cd8") return 9_450_000;
    return this.guide(value);
  }

  private argument(parts: FormulaParts, index: number): number | undefined {
    const value = parts[index];
    return value === undefined ? undefined : this.token(value);
  }

  private radians(angle: number): number {
    return (angle / 60_000) * (Math.PI / 180);
  }

  private formula(parts: FormulaParts): number | undefined {
    const [operator] = parts;
    if (operator === "val") return this.argument(parts, 1);
    if (operator === "*/") {
      const [left, multiply, divide] = [1, 2, 3].map((index) => this.argument(parts, index));
      return left === undefined || multiply === undefined || divide === undefined || divide === 0
        ? undefined
        : (left * multiply) / divide;
    }
    if (operator === "+-") {
      const [left, add, subtract] = [1, 2, 3].map((index) => this.argument(parts, index));
      return left === undefined || add === undefined || subtract === undefined
        ? undefined
        : left + add - subtract;
    }
    if (operator === "+/") {
      const [left, add, divide] = [1, 2, 3].map((index) => this.argument(parts, index));
      return left === undefined || add === undefined || divide === undefined || divide === 0
        ? undefined
        : (left + add) / divide;
    }
    if (operator === "?:") {
      const [condition, left, right] = [1, 2, 3].map((index) => this.argument(parts, index));
      return condition === undefined || left === undefined || right === undefined
        ? undefined
        : condition > 0
          ? left
          : right;
    }
    const values = parts.slice(1).map((part) => this.token(part));
    if (values.some((value) => value === undefined)) return undefined;
    const numbers = values as number[];
    switch (operator) {
      case "abs":
        return Math.abs(numbers[0]!);
      case "sqrt":
        return Math.sqrt(Math.max(numbers[0]!, 0));
      case "sin":
        return numbers[0]! * Math.sin(this.radians(numbers[1] ?? 0));
      case "cos":
        return numbers[0]! * Math.cos(this.radians(numbers[1] ?? 0));
      case "tan":
        return numbers[0]! * Math.tan(this.radians(numbers[1] ?? 0));
      case "at2":
        return Math.atan2(numbers[1]!, numbers[0]!) * 60_000;
      case "min":
        return Math.min(...numbers);
      case "max":
        return Math.max(...numbers);
      case "mod":
        return Math.sqrt(numbers.reduce((total, value) => total + value * value, 0));
      case "pin": {
        const [value, minimum, maximum] = numbers;
        return Math.max(value!, Math.min(minimum!, maximum!));
      }
      case "cat2":
        return numbers[0]! * Math.cos(Math.atan2(numbers[2]!, numbers[1]!));
      case "sat2":
        return numbers[0]! * Math.sin(Math.atan2(numbers[2]!, numbers[1]!));
      default:
        return undefined;
    }
  }
}
