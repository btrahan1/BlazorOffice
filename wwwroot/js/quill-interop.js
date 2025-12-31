window.QuillFunctions = {
    createQuill: function (element, options) {
        if (!element) return;

        // Prevent re-initialization
        if (element.__quill) return;

        var quill = new Quill(element, options || {
            theme: 'snow',
            modules: {
                toolbar: [
                    [{ 'header': [1, 2, 3, false] }],
                    ['bold', 'italic', 'underline', 'strike'],
                    [{ 'color': [] }, { 'background': [] }],
                    [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                    ['link', 'clean']
                ]
            }
        });

        element.__quill = quill;
    },

    getHTML: function (element) {
        if (!element || !element.__quill) return "";
        return element.__quill.root.innerHTML;
    },

    loadHTML: function (element, content) {
        if (!element || !element.__quill) return;
        element.__quill.root.innerHTML = content;
    },

    getText: function (element) {
        if (!element || !element.__quill) return "";
        return element.__quill.getText();
    }
};
