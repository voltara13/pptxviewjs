// External: Chart.js is provided as a global (loaded via a separate <script>).
export default (typeof globalThis !== 'undefined' ? globalThis.Chart : undefined);
