declare module "react" {
  export type ChangeEvent<T> = { target: T; currentTarget: T };
  export type SetStateAction<S> = S | ((previous: S) => S);
  export type Dispatch<T> = (value: T) => void;
  export function useState<S>(initial: S | (() => S)): [S, Dispatch<SetStateAction<S>>];
  export function useEffect(effect: () => void | (() => void), dependencies?: unknown[]): void;
  export function useMemo<T>(factory: () => T, dependencies: unknown[]): T;
  export function useRef<T>(initial: T | null): { current: T | null };
  export const StrictMode: (props: { children?: unknown }) => unknown;
}
declare module "react/jsx-runtime" {
  export namespace JSX { interface IntrinsicElements { [element: string]: unknown } }
}
declare module "react-dom/client" {
  export function createRoot(element: Element): { render(node: unknown): void };
}
