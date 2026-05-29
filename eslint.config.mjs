import nextCoreWebVitals from "eslint-config-next/core-web-vitals";
import nextTypescript from "eslint-config-next/typescript";

const eslintConfig = [
  ...nextCoreWebVitals,
  ...nextTypescript,
  {
    rules: {
      // Allow non-null assertions (common in pdf-lib / DOM APIs)
      "@typescript-eslint/no-non-null-assertion": "off",
      // Allow unused vars prefixed with _ (common pattern for destructuring)
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_" }],
      // Keep console for debugging in a client-side tool
      "no-console": "warn",
    },
  },
  {
    ignores: ["node_modules/**", ".next/**", "out/**"],
  },
];

export default eslintConfig;
