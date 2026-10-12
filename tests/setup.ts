import { hyperFormulaEngine } from '../src/core/calc';
import { chooseEngine } from '../src/core/engine';
import { nativeEngine } from '../src/engine/native';

// Sessions opened in the tests calculate with the editor's own engine, as the worker's do by default;
// ENGINE=hyperformula runs them all on HyperFormula instead.
chooseEngine(process.env.ENGINE === 'hyperformula' ? hyperFormulaEngine : nativeEngine);
