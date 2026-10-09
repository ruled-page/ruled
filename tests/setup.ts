import { hyperFormulaEngine } from '../src/core/calc';
import { chooseEngine } from '../src/core/engine';
import { nativeEngine } from '../src/engine/native';

// Sessions opened in the tests calculate with HyperFormula, as the worker's do by default;
// ENGINE=native runs them all on the editor's own engine instead.
chooseEngine(process.env.ENGINE === 'native' ? nativeEngine : hyperFormulaEngine);
