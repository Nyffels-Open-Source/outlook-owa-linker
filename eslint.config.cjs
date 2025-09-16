module.exports = [
    {
        files: ["**/*.js"],
        ignores: [
            "node_modules/**",
            "dist/**",
            "build/**"
        ],
        languageOptions: {
            ecmaVersion: "latest",
            sourceType: "module",
            globals: {
                window: "readonly",
                document: "readonly",
                navigator: "readonly",
                URL: "readonly",
                Office: "readonly"
            }
        },
        rules: {
            "no-unused-vars": "warn",
            "no-undef": "error",
            "semi": ["error", "always"],
            "quotes": ["warn", "double"],
            "no-console": "off"
        },
    }
];