window.blazorUtils = {
    createBlobUrl: async (contentStreamReference, mimeType) => {
        const arrayBuffer = await contentStreamReference.arrayBuffer();
        const blob = new Blob([arrayBuffer], { type: mimeType });
        return URL.createObjectURL(blob);
    },
    createUrlFromBlob: (blob) => {
        return URL.createObjectURL(blob);
    },
    revokeBlobUrl: (url) => {
        URL.revokeObjectURL(url);
    }
};
