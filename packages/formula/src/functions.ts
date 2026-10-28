import { formulaFunctions } from "./catalog";
import type { FormulaFunctionCatalogEntry, FormulaFunctionCategory } from "./catalog-types";
import type { FormulaEvaluationContext } from "./context";
import { FormulaValueError } from "./errors";
import { complexExtraFunctions } from "./functions/complex-extras";
import { coreExtraFunctions } from "./functions/core-extras";
import { databaseFunctions } from "./functions/database";
import { dateTimeFunctions } from "./functions/date-time";
import { dynamicArrayFunctions } from "./functions/dynamic-array";
import { dynamicExtraFunctions } from "./functions/dynamic-extras";
import { engineeringFunctions } from "./functions/engineering";
import { engineeringExtraFunctions } from "./functions/engineering-extras";
import { externalExtraFunctions } from "./functions/external-extras";
import { financialFunctions } from "./functions/financial";
import { financialExtraFunctions } from "./functions/financial-extras";
import { forecastFunctions } from "./functions/forecast-extras";
import { fourierFunctions } from "./functions/fourier-extras";
import { lookupExtraFunctions } from "./functions/lookup-extras";
import { lookupReferenceFunctions } from "./functions/lookup-reference";
import { mathFunctions } from "./functions/math";
import { statisticalFunctions } from "./functions/statistical";
import { statisticalCompatFunctions } from "./functions/statistical-compat";
import { statisticalExtraFunctions } from "./functions/statistical-extras";
import { textFunctions } from "./functions/text";
import type {
  FormulaFunctionAstEvaluator,
  FormulaFunctionImplementation,
  FormulaFunctionMetadata,
  RegisteredFunction,
} from "./functions/types";
import { type FormulaFunctionTable, define } from "./functions/utils";
import { webHostFunctions } from "./functions/web-host";
import type { FormulaValue } from "./value";
export type {
  FormulaFunctionAstEvaluator,
  FormulaFunctionImplementation,
  FormulaFunctionMetadata,
  RegisteredFunction,
} from "./functions/types";

function catalogMetadata(entry: FormulaFunctionCatalogEntry): FormulaFunctionMetadata {
  return Object.freeze({
    minArgs: entry.minArgs,
    maxArgs: Number.isFinite(entry.maxArgs) ? entry.maxArgs : "variadic",
    astOnly: entry.astOnly,
    volatile: entry.volatile,
    dynamicDependencies: entry.dynamicDependencies,
    broadcast:
      (broadcastCategories.has(entry.category) || broadcastFunctions.has(entry.name)) &&
      !nonBroadcastFunctions.has(entry.name),
    acceptsErrors: entry.acceptsErrors,
    hiddenRowSensitive: entry.name === "AGGREGATE" || entry.name === "SUBTOTAL",
    externalContext: entry.externalContext,
  });
}

const broadcastCategories = new Set<FormulaFunctionCategory>([
  "dateTime",
  "engineering",
  "financial",
  "math",
  "information",
  "text",
]);

const broadcastFunctions = new Set([
  "BETA.DIST",
  "BETA.INV",
  "BETADIST",
  "BETAINV",
  "BINOM.DIST",
  "BINOM.INV",
  "BINOMDIST",
  "CHIDIST",
  "CHIINV",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CONFIDENCE",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "EXPON.DIST",
  "EXPONDIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "FDIST",
  "FINV",
  "FISHER",
  "FISHERINV",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMADIST",
  "GAMMAINV",
  "GAMMALN.PRECISE",
  "LOGINV",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "LOGNORMDIST",
  "NORM.DIST",
  "NORMDIST",
  "NORM.INV",
  "NORMINV",
  "NORM.S.DIST",
  "NORMSDIST",
  "NORM.S.INV",
  "NORMSINV",
  "POISSON",
  "POISSON.DIST",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "TDIST",
  "TINV",
  "WEIBULL",
  "WEIBULL.DIST",
]);

const nonBroadcastFunctions = new Set([
  "AGGREGATE",
  "ARRAYTOTEXT",
  "CONCAT",
  "GCD",
  "IMPORTRANGE",
  "IMPRODUCT",
  "IMSUM",
  "LCM",
  "MMULT",
  "MDETERM",
  "MUNIT",
  "MINVERSE",
  "NETWORKDAYS",
  "NETWORKDAYS.INTL",
  "PRODUCT",
  "SINGLE",
  "STRINGS",
  "STOCKHISTORY",
  "SUBTOTAL",
  "SUM",
  "SUMIF",
  "SUMIFS",
  "SUMSQ",
  "SUMX2MY2",
  "SUMX2PY2",
  "SUMXMY2",
  "SUMPRODUCT",
  "TEXTJOIN",
  "VALUES",
  "WEEKNUM",
  "WORKDAY",
  "WORKDAY.INTL",
  "IRR",
  "NPV",
  "XIRR",
  "XNPV",
]);

const catalogByName = new Map(formulaFunctions.map((entry) => [entry.name, entry]));

const duplicateCatalogNames = formulaFunctions.filter(
  (entry, index) =>
    formulaFunctions.findIndex((candidate) => candidate.name === entry.name) !== index,
);
if (duplicateCatalogNames.length > 0)
  throw new Error(
    `duplicate formula catalog names: ${duplicateCatalogNames.map((entry) => entry.name).join(", ")}`,
  );

function registerFunctions(
  table: FormulaFunctionTable,
  options: { override?: boolean } = {},
): void {
  for (const [name, implementation] of table) {
    const entry = catalogByName.get(name);
    if (!entry) throw new Error(`uncataloged formula function: ${name}`);
    if (!options.override && registry.has(name))
      throw new Error(`duplicate formula function registration: ${name}`);
    const registered = define(catalogMetadata(entry), implementation);
    registry.set(name, registered);
  }
}

type MutableRegisteredFunction = FormulaFunctionImplementation & {
  metadata: FormulaFunctionMetadata;
  evaluateAst?: FormulaFunctionAstEvaluator;
};

const registry = new Map<string, MutableRegisteredFunction>();

registerFunctions(mathFunctions);
registerFunctions(dateTimeFunctions);
registerFunctions(textFunctions);
registerFunctions(statisticalFunctions);
registerFunctions(dynamicArrayFunctions);
registerFunctions(lookupReferenceFunctions);
registerFunctions(financialFunctions);
registerFunctions(engineeringFunctions);
registerFunctions(databaseFunctions);
registerFunctions(webHostFunctions);
registerFunctions(coreExtraFunctions);
registerFunctions(engineeringExtraFunctions);
registerFunctions(financialExtraFunctions);
registerFunctions(statisticalExtraFunctions, { override: true });
registerFunctions(forecastFunctions, { override: true });
registerFunctions(statisticalCompatFunctions);
registerFunctions(dynamicExtraFunctions);
registerFunctions(lookupExtraFunctions);
registerFunctions(complexExtraFunctions);
registerFunctions(externalExtraFunctions);
registerFunctions(fourierFunctions);

for (const entry of formulaFunctions.filter((candidate) => candidate.astOnly)) {
  const metadata = catalogMetadata(entry);
  registry.set(
    entry.name,
    define(metadata, () => {
      throw new FormulaValueError("#VALUE!");
    }),
  );
}

const missingFunctions = formulaFunctions
  .filter((entry) => !entry.astOnly && !registry.has(entry.name))
  .map((entry) => entry.name);
if (missingFunctions.length > 0) {
  throw new Error(
    `unimplemented formula functions (${missingFunctions.length}): ${missingFunctions.join(", ")}`,
  );
}

function normalizeFunctionName(name: string): string {
  return name
    .toUpperCase()
    .replace(/^(?:(?:_XLFN|_XLWS|ORG\.OPENOFFICE(?:\.ORG)?|COM\.MICROSOFT)\.)+/u, "");
}

const publicFunctions = new WeakMap<RegisteredFunction, RegisteredFunction>();

export function getFunction(name: string): RegisteredFunction | undefined {
  const implementation = registry.get(name) ?? registry.get(normalizeFunctionName(name));
  if (!implementation) return undefined;
  let view = publicFunctions.get(implementation);
  if (!view) {
    view = Object.assign(
      (args: FormulaValue[], context: FormulaEvaluationContext) => implementation(args, context),
      {
        metadata: implementation.metadata,
        evaluateAst: implementation.evaluateAst,
      },
    );
    Object.defineProperties(view, {
      metadata: {
        enumerable: true,
        get(): FormulaFunctionMetadata {
          return implementation.metadata;
        },
      },
      evaluateAst: {
        enumerable: true,
        get(): FormulaFunctionAstEvaluator | undefined {
          return implementation.evaluateAst;
        },
      },
    });
    Object.freeze(view);
    publicFunctions.set(implementation, view);
  }
  return view;
}

export function listFunctionNames(): string[] {
  return [...registry.keys()].sort((left, right) => left.localeCompare(right));
}

export function setFunctionAstEvaluator(
  name: string,
  evaluateAst: FormulaFunctionAstEvaluator,
): void {
  const normalized = normalizeFunctionName(name);
  const implementation = registry.get(normalized);
  if (!implementation) throw new Error(`unregistered AST formula function: ${name}`);
  implementation.evaluateAst = evaluateAst;
}

export function clearFunctionAstEvaluator(name: string): void {
  const normalized = normalizeFunctionName(name);
  const implementation = registry.get(normalized);
  if (!implementation) throw new Error(`unregistered AST formula function: ${name}`);
  delete implementation.evaluateAst;
}

export function getFunctionMetadata(name: string): FormulaFunctionMetadata | undefined {
  return getFunction(name)?.metadata;
}
