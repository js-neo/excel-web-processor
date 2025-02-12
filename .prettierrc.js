export default {
    trailingComma: "none",
    tabWidth: 4,
    semi: true,
    printWidth: 80,
    parser: "babel",
    overrides: [
        {
            files: "*.js",
            options: {
                parser: "babel"
            }
        }
    ]
};

