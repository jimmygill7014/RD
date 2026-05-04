import { createContext, useCallback, useContext, useEffect, useRef, useState } from 'react';
import { ensureUidsDeep } from './uid.js';

export const STORE_KEY = 'pfa-intake-v2';
const AUTOSAVE_DEBOUNCE_MS = 800;
const SAVED_FLASH_MS = 2000;

const StoreContext = createContext(null);

function readFromStorage() {
  try {
    return ensureUidsDeep(JSON.parse(localStorage.getItem(STORE_KEY) || '{}'));
  } catch {
    return {};
  }
}

function writeToStorage(data) {
  localStorage.setItem(STORE_KEY, JSON.stringify(data));
}

export function StoreProvider({ children }) {
  const [data, setData] = useState(readFromStorage);
  const [autosaveStatus, setAutosaveStatus] = useState('idle');

  const autosaveTimer = useRef(null);
  const flashTimer = useRef(null);
  const isInitialMount = useRef(true);

  useEffect(() => {
    if (isInitialMount.current) {
      isInitialMount.current = false;
      return;
    }
    setAutosaveStatus('saving');
    clearTimeout(autosaveTimer.current);
    clearTimeout(flashTimer.current);
    autosaveTimer.current = setTimeout(() => {
      try {
        writeToStorage(data);
        setAutosaveStatus('saved');
        flashTimer.current = setTimeout(() => setAutosaveStatus('idle'), SAVED_FLASH_MS);
      } catch (err) {
        console.error('Autosave failed', err);
        setAutosaveStatus('idle');
      }
    }, AUTOSAVE_DEBOUNCE_MS);
    return () => clearTimeout(autosaveTimer.current);
  }, [data]);

  useEffect(() => {
    return () => {
      clearTimeout(autosaveTimer.current);
      clearTimeout(flashTimer.current);
    };
  }, []);

  const update = useCallback((path, value) => {
    setData(prev => setByPath(prev, path, value));
  }, []);

  const replace = useCallback(newData => {
    setData(newData ?? {});
  }, []);

  const reset = useCallback(() => {
    clearTimeout(autosaveTimer.current);
    clearTimeout(flashTimer.current);
    localStorage.removeItem(STORE_KEY);
    setData({});
    setAutosaveStatus('idle');
  }, []);

  const saveNow = useCallback(() => {
    clearTimeout(autosaveTimer.current);
    writeToStorage(data);
    setAutosaveStatus('saved');
    clearTimeout(flashTimer.current);
    flashTimer.current = setTimeout(() => setAutosaveStatus('idle'), SAVED_FLASH_MS);
  }, [data]);

  const stripKeysMatching = useCallback(predicate => {
    setData(prev => stripDeep(prev, predicate));
  }, []);

  const value = { data, update, replace, reset, saveNow, stripKeysMatching, autosaveStatus };
  return <StoreContext.Provider value={value}>{children}</StoreContext.Provider>;
}

export function useStore() {
  const ctx = useContext(StoreContext);
  if (!ctx) throw new Error('useStore must be used inside <StoreProvider>');
  return ctx;
}

function stripDeep(obj, predicate) {
  if (Array.isArray(obj)) {
    return obj.map(item =>
      item && typeof item === 'object' ? stripDeep(item, predicate) : item
    );
  }
  if (!obj || typeof obj !== 'object') return obj;
  const next = {};
  for (const key of Object.keys(obj)) {
    if (predicate(key)) continue;
    const v = obj[key];
    next[key] = v && typeof v === 'object' ? stripDeep(v, predicate) : v;
  }
  return next;
}

function setByPath(obj, path, value) {
  const keys = String(path).split('.');
  const next = Array.isArray(obj) ? [...obj] : { ...obj };
  let cursor = next;
  for (let i = 0; i < keys.length - 1; i++) {
    const key = keys[i];
    const existing = cursor[key];
    cursor[key] = existing && typeof existing === 'object' ? { ...existing } : {};
    cursor = cursor[key];
  }
  cursor[keys[keys.length - 1]] = value;
  return next;
}
