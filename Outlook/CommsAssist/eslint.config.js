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
            "no-unused-vars": "warn",
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
