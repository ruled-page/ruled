import { chooseEngine } from '../src/core/engine';
import { nativeEngine } from '@ruled/formula/native';

// Sessions opened in the tests calculate with the editor's engine, as the worker's do.
chooseEngine(nativeEngine);
