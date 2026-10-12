import { chooseEngine } from '../src/core/engine';
import { nativeEngine } from '../src/engine/native';

// Sessions opened in the tests calculate with the editor's engine, as the worker's do.
chooseEngine(nativeEngine);
