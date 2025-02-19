export const sortKeys = (a, b) => {
    const splitRegex = /(\d+)|([^0-9]+)/g;

    const getParts = (str) =>
        str.match(splitRegex)?.filter((x) => x !== undefined && x !== "") || [];

    const aParts = getParts(a);
    const bParts = getParts(b);

    const maxLength = Math.max(aParts.length, bParts.length);

    for (let i = 0; i < maxLength; i++) {
        const aPart = aParts[i] || "";
        const bPart = bParts[i] || "";

        const aIsNum = /^\d+$/.test(aPart);
        const bIsNum = /^\d+$/.test(bPart);

        if (aIsNum && bIsNum) {
            const aNum = parseInt(aPart, 10);
            const bNum = parseInt(bPart, 10);
            if (aNum !== bNum) return aNum - bNum;
        } else {
            const diff = aPart.localeCompare(bPart, undefined, {
                sensitivity: "base"
            });
            if (diff !== 0) return diff;
        }
    }
    return 0;
};
