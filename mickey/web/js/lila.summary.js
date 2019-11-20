// s1: summary home
var s1;
var SummaryHome = {
    settings: {
        workspace: $("#workspace"),
        function: $("#function"),
        btnCheck: $("#btn-check"),
        table: $("#summary-table"),
        progressBar: $("#progress-bar"),
        alterBox: $("#alert-box")
    },

    init: function() {
        s1 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        s1.workspace.on("change", async function() {
            // Update
            info = {
                name: s1.workspace.val()
            };
            await eel.update_workspace(info, "update");

            s1.table.html("");
            s1.alterBox.show();
        });

        s1.btnCheck.on("click", async function() {
            s1.btnCheck.attr("disabled", true);
            s1.alterBox.hide();
            s1.table.html("");

            var data = await eel.get_workspace_summary(true)();
            SummaryHome.render(data);
            SummaryHome.formatTable();

            s1.btnCheck.attr("disabled", false);
        });

        s1.table.on("click", ".sync", async function() {
            var testlog = $(this).attr("path");
            var package = s1.workspace.find("option:selected").attr("package");

            $("#s2-warning-tbl tbody").remove();
            SummaryModal.invoke();

            var data = await eel.check_testlog(testlog, package)();
            if (data.warninglist != undefined) {
                $("#s2-warning-tbl").append(data.warninglist);
            }
            s2.wait.hide();
        });
    },

    render: function(data) {
        s1.workspace.selectpicker({ noneSelectedText: "Please add Workspace" });

        // Render workspace
        if (data.workspace != undefined) s1.workspace.html(data.workspace);
        if (data.lastWsp != undefined) s1.workspace.val(data.lastWsp);

        // Render table
        if (data.table != undefined) s1.table.html(data.table);

        // Refresh select
        s1.workspace.selectpicker("refresh");
    },

    formatTable: function() {
        $("#summary-table").removeClass();
        $("#summary-table").addClass(
            "table m-t-10 toggle-arrow-tiny color-table info-table"
        );
        $("#summary-table")
            .footable()
            .on("footable_row_expanded", function(e) {
                $("#summary-table tbody tr.footable-detail-show")
                    .not(e.row)
                    .each(function() {
                        $("#summary-table")
                            .data("footable")
                            .toggleDetail(this);
                    });
            });
    },

    updateProgressBar: function(percent) {
        var style = "width: " + percent + "%; height:14px;";
        s1.progressBar.attr("style", style);
    }
};

// s2: summary table
var s2;
var SummaryModal = {
    settings: {
        modal: $("#modal-s2"),
        table: $("#s2-warning-tbl"),
        wait: $("#s2-wait")
    },

    init: function() {
        s2 = this.settings;
    },

    invoke: function() {
        s2.wait.show();
        s2.modal.modal();
    }
};

// Initial
SummaryHome.init();
SummaryModal.init();

// Update process
eel.expose(updateProgress);
function updateProgress(data) {
    if (data.percent != undefined) {
        SummaryHome.updateProgressBar(data.percent);
        if (data.percent == 100) {
            $("#progress-bar").hide();
        }
    }

    SummaryHome.render(data);
}

// Onload
window.onload = async function(e) {
    await translate();
    var data = await eel.get_workspace_summary()();
    SummaryHome.render(data);
    $(".preloader").fadeOut();
};
