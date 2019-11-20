var Utils = {
    selectFolder: async function(element) {
        var folder = await eel.select_folder()();
        var tag = element.prop("tagName").toLowerCase();
        switch (tag) {
            case "input":
                element.val(folder);
                break;
            default:
                element.text(folder);
        }
    },

    selectFile: async function(element) {
        var file = await eel.select_file("excel")();
        var tag = element.prop("tagName").toLowerCase();
        switch (tag) {
            case "input":
                element.val(file);
                break;
            default:
                element.text(file);
        }
    }
};

// Change language
$("#language").click(async function() {
    await eel.change_language($("#language").text())();
    location.reload();
});

async function translate(id = "") {
    if (window.dict == undefined) {
        window.dict = await eel.get_language_text()();
    }

    $(id + " .lang").each(function() {
        var word = $(this).text().trim();
        if (dict[word] != undefined) {
            $(this).text(dict[word]);
        }
    });
}

async function updateMessage() {
    var data = await eel.get_messages()();

    if (data.html != undefined) {
        $("#notification").show();
        $("#messages").html(data.html);
    }
}

$(document).ready(async function() {
    await updateMessage();
    await eel.sync_package_data()();
});
