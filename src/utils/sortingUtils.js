export const sortKeys = (a, b) => {
    const splitRegex = /(\d+([.,]\d+)?)|([^0-9]+)/g;

    const getParts = (str) =>
        str.match(splitRegex)?.filter((x) => x !== undefined && x !== "") || [];

    const aParts = getParts(a);
    const bParts = getParts(b);

    for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
        const aPart = aParts[i] || "";
        const bPart = bParts[i] || "";

        const aIsNum = /^\d+([.,]\d+)?$/.test(aPart);
        const bIsNum = /^\d+([.,]\d+)?$/.test(bPart);

        if (aIsNum && bIsNum) {
            const aNum = parseFloat(aPart.replace(",", "."));
            const bNum = parseFloat(bPart.replace(",", "."));
            if (aNum !== bNum) return aNum - bNum;
        } else if (aIsNum !== bIsNum) {
            return aIsNum ? -1 : 1;
        } else {
            const diff = aPart.localeCompare(bPart, undefined, {
                sensitivity: "base"
            });
            if (diff !== 0) return diff;
        }
    }
    return 0;
};
