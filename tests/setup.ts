import { hyperFormulaEngine } from '../src/core/calc';
import { chooseEngine } from '../src/core/engine';

// Sessions opened in the tests calculate with HyperFormula, as the worker's do by default.
chooseEngine(hyperFormulaEngine);
