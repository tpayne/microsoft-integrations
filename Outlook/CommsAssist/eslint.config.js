import eslintJs from "@eslint/js";
import officeAddins from "eslint-plugin-office-addins";

export default [
    {
        files: ["src/**/*.js"],
        plugins: {
            "office-addins": officeAddins,
        },
        rules: {
            ...eslintJs.configs.recommended.rules,
            ...officeAddins.configs.recommended.rules,
            "no-unused-vars": [
                "warn",
                {
                    // Default options for other unused variables (func args, local vars)
                    vars: "all",
                    args: "after-used",
                    ignoreRestSiblings: true,
                    // Fix 1: Treat caught errors as a separate group
                    caughtErrors: "all",
                    // Fix 2: Ignore any caught error variable named 'e' or 'openErr'
                    caughtErrorsIgnorePattern: "^(e|openErr)$",
                }
            ],
            "no-undef": "warn",
        },
        languageOptions: {
            globals: {
                Office: "readonly",
                Word: "readonly",
                Excel: "readonly",
                PowerPoint: "readonly",
                window: "readonly",
                document: "readonly",
                console: "readonly",
                fetch: "readonly",
                setTimeout: "readonly",
            },
        },
    },
];