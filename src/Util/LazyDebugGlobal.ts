/**
 * Defines `scope[name]` as a value that is computed by `compute` the first time it is read,
 * then kept as a plain writable data property. Assigning `scope[name]` before the first read
 * replaces it like a plain property, without computing.
 *
 * Used for debug state that is exported on every render but only read by debug tooling:
 * the export captures its render-time inputs, and the expensive part runs on demand.
 */
export function defineLazyDebugGlobal(scope: any, name: string, compute: () => any): void {
    const store: (value: any) => void = (value: any): void => {
        Object.defineProperty(scope, name, { configurable: true, enumerable: true, writable: true, value });
    };
    Object.defineProperty(scope, name, {
        configurable: true,
        enumerable: true,
        get: (): any => {
            const value: any = compute();
            store(value);
            return value;
        },
        set: (value: any): void => {
            store(value);
        },
    });
}
