// p1: package manager home
var p1;
var PackageHome = {
    settings: {
        table: $("#package-table"),
        btnNewPackage: $("#btn-new-package")
    },

    init: function() {
        p1 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        p1.btnNewPackage.on("click", function() {
            PackageModal.invoke();
        });

        p1.table.on("click", ".delete", function() {
            //get the footable object
            var footable = p1.table.data("footable");

            //get the row we are wanting to delete
            var row = $(this).parents("tr:first");

            var cell = $(this)
                .parents("tr:first")
                .find("td:first");
            var info = {
                name: cell.text()
            };

            // delete package
            eel.update_package(info, "delete");

            //delete the row
            footable.removeRow(row);
        });

        p1.table.on("click", ".edit", async function() {
            var cell = $(this)
                .parents("tr:first")
                .find("td:first");
            var name = cell.text();

            var info = await eel.get_package_info(name)();

            PackageModal.invoke();
            PackageModal.render(info);
        });

        p1.table.on("click", ".sync", async function() {
            var cell = $(this)
                .parents("tr:first")
                .find("td:first");
            var name = cell.text();

            $(".preloader").fadeIn();
            await eel.update_package_data(name)();
            $(".preloader").fadeOut();
        });
    },

    addRow: function(data) {
        $("#package-table tbody").append(
            $("<tr>")
                .append($("<td>").append(data.name))
                .append($("<td>").append(data.jira))
                .append(
                    $("<td>").append(
                        '<button type="button" class="btn btn-sm btn-icon btn-pure btn-outline edit" data-toggle="tooltip" data-original-title="Edit"><i class="mdi mdi-settings" aria-hidden="true"></i></button>' +
                            '<button type="button" class="btn btn-sm btn-icon btn-pure btn-outline delete" data-toggle="tooltip" data-original-title="Edit"><i class="mdi mdi-delete" aria-hidden="true"></i></button>' +
                            '<button type="button" class="btn btn-sm btn-icon btn-pure btn-outline sync" data-toggle="tooltip" data-original-title="Edit"><i class="mdi mdi-sync" aria-hidden="true"></i></button>'
                    )
                )
                .append($("<td>").append(data.summary.xlsx))
                .append($("<td>").append(data.summary.sheet))
                .append($("<td>").append(data.summary.begin))
                .append($("<td>").append(data.summary.end))
                .append($("<td>").append(data.report.xlsx))
                .append($("<td>").append(data.report.sheet))
                .append($("<td>").append(data.report.begin))
                .append($("<td>").append(data.report.end))
                .append($("<td>").append(data.source))
        );
    },

    formatTable: function() {
        $("#package-table")
            .footable()
            .on("footable_row_expanded", function(e) {
                $("#package-table tbody tr.footable-detail-show")
                    .not(e.row)
                    .each(function() {
                        $("#package-table")
                            .data("footable")
                            .toggleDetail(this);
                    });
            });
    },

    buildTable: async function() {
        // Build table
        var data = await eel.get_package_info()();
        for (var name in data) {
            PackageHome.addRow(data[name]);
        }

        // Format table
        PackageHome.formatTable();
    }
};

// p2: package manager modal
var p2;
var PackageModal = {
    settings: {
        form: $("#p2-form"),
        modal: $("#modal-p2"),
        name: $("#p2-name"),
        jira: $("#p2-jira"),
        summary: $("#p2-summary"),
        summarySheet: $("#p2-summary-sheet"),
        summaryBegin: $("#p2-summary-begin"),
        summaryEnd: $("#p2-summary-end"),
        report: $("#p2-report"),
        reportSheet: $("#p2-report-sheet"),
        reportBegin: $("#p2-report-begin"),
        reportEnd: $("#p2-report-end"),
        source: $("#p2-source"),
        btnSave: $("#p2-btn-save"),
        warningBox: $("#p2-warning")
    },

    init: function() {
        p2 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        p2.btnSave.on("click", async function() {
            info = {
                name: p2.name.val(),
                summary: {
                    xlsx: p2.summary.val(),
                    sheet: p2.summarySheet.val(),
                    begin: parseInt(p2.summaryBegin.val()),
                    end: parseInt(p2.summaryEnd.val())
                },
                report: {
                    xlsx: p2.report.val(),
                    sheet: p2.reportSheet.val(),
                    begin: parseInt(p2.reportBegin.val()),
                    end: parseInt(p2.reportEnd.val())
                },
                jira: p2.jira.val(),
                source: p2.source.val()
            };

            if (
                p2.name.val().trim() == "" ||
                p2.summary.val().trim() == "" ||
                p2.summarySheet.val().trim() == "" ||
                p2.jira.val().trim() == "" ||
                p2.source.val().trim() == "" ||
                p2.source.val().trim() == "."
            ) {
                p2.warningBox.show();
            } else {
                PackageModal.updatePackage("update");
                p2.modal.modal("toggle");
                window.location.reload();
            }
        });

        p2.source.on("click", function() {
            Utils.selectFolder(p2.source);
        });
    },

    invoke: function() {
        p2.form[0].reset();
        p2.modal.modal();
        p2.name.prop("disabled", false);
        p2.warningBox.hide();
    },

    render: async function(data) {
        if (data.name != undefined) p2.name.val(data.name);
        if (data.jira != undefined) p2.jira.val(data.jira);
        if (data.source != undefined) p2.source.val(data.source);

        if (data.summary != undefined) {
            p2.summary.val(data.summary.xlsx);
            p2.summarySheet.val(data.summary.sheet);
            p2.summaryBegin.val(data.summary.begin);
            p2.summaryEnd.val(data.summary.end);
        }

        if (data.report != undefined) {
            p2.report.val(data.report.xlsx);
            p2.reportSheet.val(data.report.sheet);
            p2.reportBegin.val(data.report.begin);
            p2.reportEnd.val(data.report.end);
        }

        p2.name.prop("disabled", true);
    },

    updatePackage: async function(action) {
        info = {
            name: p2.name.val(),
            summary: {
                xlsx: p2.summary.val(),
                sheet: p2.summarySheet.val(),
                begin: parseInt(p2.summaryBegin.val()),
                end: parseInt(p2.summaryEnd.val())
            },
            report: {
                xlsx: p2.report.val(),
                sheet: p2.reportSheet.val(),
                begin: parseInt(p2.reportBegin.val()),
                end: parseInt(p2.reportEnd.val())
            },
            jira: p2.jira.val(),
            source: p2.source.val()
        };

        await eel.update_package(info, action)();
    }
};

// Init
PackageHome.init();
PackageModal.init();

// Onload
window.onload = async function(e) {
    await translate();
    PackageHome.buildTable();
    $(".preloader").fadeOut();
};
