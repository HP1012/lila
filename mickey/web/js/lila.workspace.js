// a1: workspace home
var a1;
var WorkspaceHome = {
    settings: {
        workspace: $("#workspace"),
        function: $("#function"),
        btnNew: $("#btn-new-wsp"),
        btnDelete: $("#btn-delete-wsp"),
        btnEdit: $("#btn-edit-wsp"),
        btnCheck: $("#btn-check"),
        btnReport: $("#btn-report")
    },

    init: function() {
        a1 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        a1.btnNew.on("click", function() {
            WorkspaceModal.invokeNew();
        });

        a1.btnDelete.on("click", function() {
            WorkspaceModal.invokeDelete();
        });

        a1.btnEdit.on("click", function() {
            WorkspaceModal.invokeEdit();
        });

        a1.workspace.on("change", async function() {
            // Update
            info = {
                name: a1.workspace.val()
            };
            await eel.update_workspace(info, "update");

            // Rebuild
            WorkspaceHome.buildWorkspace(3);
        });

        a1.function.on("change", async function() {
            WorkspaceHome.check();
        });

        a1.btnCheck.on("click", async function() {
            WorkspaceHome.check();
        });

        a1.btnReport.on("click", function() {
            ReportModal.invoke();
        });
    },

    render: function(data) {
        a1.workspace.selectpicker({ noneSelectedText: "Please add Workspace" });
        a1.function.selectpicker({
            noneSelectedText: "Please select Function"
        });

        // Render workspace
        if (data.workspace != undefined) a1.workspace.html(data.workspace);
        if (data.function != undefined) a1.function.html(data.function);
        if (data.pkgOptions != undefined) a2.package.html(data.pkgOptions);

        if (data.lastWsp != undefined) a1.workspace.val(data.lastWsp);
        if (data.lastFunc != undefined) a1.function.val(data.lastFunc);

        if (data.checklist != undefined)
            $("#checklist-tbl").append(data.checklist);
        if (data.warninglist != undefined)
            $("#warning-tbl").append(data.warninglist);

        // Refresh select
        a1.workspace.selectpicker("refresh");
        a1.function.selectpicker("refresh");
        a2.package.selectpicker("refresh");
    },

    clean: function(level) {
        if (level == 0) {
            return;
        }

        // When change function
        if (level > 1) {
            // Clean checklist and warning list
            $("#checklist-tbl tbody").remove();
            $("#warning-tbl tbody").remove();

            // Clean coverage report table
            $("#coverage-report [data-name]").each(function() {
                $(this).text("");
            });

            // Clean function info table
            $("#function-info [data-name]").each(function() {
                $(this).text("");
            });
        }

        // When change workspace
        if (level > 2) {
            a1.function.html("");
        }

        // Add/delete new workspace
        if (level > 3) {
            a1.workspace.html("");
        }
    },

    check: async function() {
        // Update
        var info = {
            name: a1.workspace.val(),
            function: a1.function.find("option:selected").attr("path")
        };
        await eel.update_workspace(info, "update");

        // Rebuild
        WorkspaceHome.buildWorkspace(2);
    },

    // Render Coverage Report
    renderCoverageReport: function(data) {
        $("#coverage-report [data-name]").each(function() {
            var key = $(this).attr("data-name");
            if (data[key] != undefined) $(this).text(data[key]);
        });
    },

    // Render Function Info
    renderFunctionInfo: function(data) {
        $("#function-info [data-name]").each(function() {
            var key = $(this).attr("data-name");
            if (data[key] != undefined) $(this).text(data[key]);
        });
    },

    updateFunctionInfo: async function() {
        // Disable button
        a1.btnCheck.attr("disabled", true);
        a1.btnReport.attr("disabled", true);

        var testlog = a1.function.find("option:selected").attr("path");
        var package = a1.workspace.find("option:selected").attr("package");

        if (testlog != undefined) {
            // Show test coverage info
            var data = await eel.get_coverage_report(testlog)();
            WorkspaceHome.renderCoverageReport(data);

            // Show function info from summary file
            data = await eel.get_function_info(testlog, package)();
            WorkspaceHome.renderFunctionInfo(data);

            // Check result
            data = await eel.check_testlog(testlog, package)();
            this.render(data);
        }

        // Enable button
        a1.btnCheck.attr("disabled", false);
        a1.btnReport.attr("disabled", false);
    },

    buildWorkspace: async function(level) {
        // Clean workspace
        WorkspaceHome.clean(level);

        var data = await eel.get_workspace_data()();
        WorkspaceHome.render(data);

        WorkspaceHome.updateFunctionInfo();
    }
};

// a2: workspace manager modal
var a2;
var WorkspaceModal = {
    settings: {
        modal: $("#modal-a2"),
        form: $("#a2-form"),
        path: $("#a2-path"),
        package: $("#a2-package"),
        name: $("#a2-name"),
        deliver: $("#a2-deliver"),
        btnSave: $("#a2-btn-save"),
        btnDelete: $("#a2-btn-delete"),
        btnPath: $("#a2-btn-path"),
        btnDeliver: $("#a2-btn-deliver"),
        warningBox: $("#a2-warning")
    },

    init: function() {
        a2 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        a2.btnPath.on("click", async function() {
            Utils.selectFolder(a2.path);
        });

        a2.btnDeliver.on("click", function() {
            Utils.selectFolder(a2.deliver);
        });

        a2.btnSave.on("click", async function() {
            if (
                a2.path.text().trim() == "" ||
                a2.path.text().trim() == "." ||
                a2.deliver.text().trim() == "" ||
                a2.deliver.text().trim() == "." ||
                a2.name.val().trim() == "" ||
                a2.package.val().trim() == ""
            ) {
                a2.warningBox.show();
            } else {
                WorkspaceModal.updateWorkspace("update");
                WorkspaceHome.buildWorkspace(4);
                a2.modal.modal("toggle");
            }
        });

        a2.btnDelete.on("click", async function() {
            WorkspaceModal.updateWorkspace("delete");
            WorkspaceHome.buildWorkspace(4);
            a2.modal.modal("toggle");
        });
    },

    invoke: function(render) {
        a2.modal.modal();
        a2.form[0].reset();
        a2.path.text("");
        a2.deliver.text("");
        a2.warningBox.hide();

        // Edit and delete
        if (render == true) {
            var workspace = a1.workspace.val();
            eel.get_workspace_info(workspace)(WorkspaceModal.render);
        }
    },

    render: function(data) {
        a2.package.selectpicker({
            noneSelectedText: "Please select Package"
        });

        if (data.path != undefined) a2.path.text(data.path);
        if (data.deliver != undefined) a2.deliver.text(data.deliver);
        if (data.name != undefined) a2.name.val(data.name);
        if (data.pkgOptions != undefined) a2.package.html(data.pkgOptions);
        if (data.package != undefined) a2.package.val(data.package);

        a2.package.selectpicker("refresh");
    },

    invokeNew: function() {
        WorkspaceModal.invoke(false);
        a2.btnDelete.hide();
        a2.btnSave.show();
        a2.name.prop("disabled", false);
    },

    invokeDelete: function() {
        WorkspaceModal.invoke(true);
        a2.btnDelete.show();
        a2.btnSave.hide();
        a2.name.prop("disabled", true);
    },

    invokeEdit: function() {
        WorkspaceModal.invoke(true);
        a2.btnDelete.hide();
        a2.btnSave.show();
        a2.name.prop("disabled", true);
    },

    updateWorkspace: async function(action) {
        var info = {
            path: a2.path.text(),
            deliver: a2.deliver.text(),
            name: a2.name.val(),
            package: a2.package.val()
        };

        await eel.update_workspace(info, action)();
    }
};

// a4: report manager modal
var a4;
var ReportModal = {
    settings: {
        modal: $("#modal-a4"),
        form: $("#a4-form"),
        target: $("#a4-target"),
        spec: $("#a4-spec"),
        template: $("#a4-template"),
        issue: $("#a4-issue"),
        funcNo: $("#a4-func-no"),
        btnUpdate: $("#a4-btn-update"),
        btnCreateNew: $("#a4-btn-create-new"),
        btnTarget: $("#a4-btn-target"),
        btnSpec: $("#a4-btn-spec")
    },

    init: function() {
        a4 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        a4.btnTarget.on("click", function() {
            Utils.selectFolder(a4.target);
        });

        a4.btnSpec.on("click", function() {
            Utils.selectFolder(a4.spec);
        });

        a4.template.on("click", async function() {
            await Utils.selectFile(a4.template);
        });

        a4.btnCreateNew.on("click", async function() {
            var package = a1.workspace.find("option:selected").attr("package");
            var info = {
                testlog: a1.function.find("option:selected").attr("path"),
                target: a4.target.text(),
                spec: a4.spec.text(),
                template: a4.template.val(),
                issue: a4.issue.val(),
                func_no: a4.funcNo.val(),
                package: package
            };

            a4.modal.modal("toggle");

            a5.progressBar.addClass("active progress-bar-striped");
            ReportProgressModal.invoke();
            a5.btnClose.hide();

            await eel.deliver_report(info)();
            a5.progressBar.removeClass("active progress-bar-striped");
            a5.btnClose.show();
        });

        a4.btnUpdate.on("click", async function() {
            var package = a1.workspace.find("option:selected").attr("package");
            var info = {
                testlog: a1.function.find("option:selected").attr("path"),
                target: a4.target.text(),
                spec: a4.spec.text(),
                template: a4.template.val(),
                issue: a4.issue.val(),
                func_no: a4.funcNo.val(),
                package: package
            };

            a4.modal.modal("toggle");

            a5.progressBar.addClass("active progress-bar-striped");
            ReportProgressModal.invoke();
            a5.btnClose.hide();

            await eel.deliver_report_update(info)();
            a5.progressBar.removeClass("active progress-bar-striped");
            a5.btnClose.show();
        });

    },

    invoke: async function() {
        a4.form[0].reset();

        // Get report info
        var testlog = a1.function.find("option:selected").attr("path");
        var package = a1.workspace.find("option:selected").attr("package");
        var workspace = a1.workspace.val();
        await eel.get_report_info(testlog, workspace)(ReportModal.render);

        a4.modal.modal();
    },

    render: function(data) {
        if (data.dirTarget != undefined) a4.target.text(data.dirTarget);
        if (data.dirSpec != undefined) a4.spec.text(data.dirSpec);

        if (data.func_no != undefined) {
            a4.funcNo.val(data.func_no);
            a4.funcNo.prop("disabled", true);
        } else {
            a4.funcNo.prop("disabled", false);
        }
    }
};

// a5: report progress modal
var a5;
var ReportProgressModal = {
    settings: {
        modal: $("#modal-a5"),
        body: $("#a5-body"),
        btnClose: $("#a5-btn-close"),
        progressBar: $("#a5-progress-bar")
    },

    init: function() {
        a5 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {},

    invoke: function() {
        a5.body.html("");
        a5.modal.modal({
            backdrop: "static",
            keyboard: false
        });
    },

    updateProgressBar: function(percent) {
        var style = "width: " + percent + "%; height:14px;";
        a5.progressBar.attr("style", style);
    }
};

// a6: auth modal
var a6;
var AuthModal = {
    settings: {
        modal: $("#modal-a6"),
        form: $("#a6-form"),
        warning: $("#a6-warning"),
        username: $("#a6-username"),
        password: $("#a6-password"),
        btnSave: $("#a6-btn-save")
    },

    init: function() {
        a6 = this.settings;
        this.bindUIActions();
    },

    bindUIActions: function() {
        a6.btnSave.on("click", async function() {
            var username = a6.username.val();
            var password = a6.password.val();

            var status = await eel.update_auth(username, password)();
            if (status == 401) {
                a6.warning.html("Incorrect username/password");
            } else {
                a6.modal.modal("toggle");
            }
        });
    },

    invoke: function() {
        a6.form[0].reset();
        a6.warning.html("");
        a6.modal.modal({
            backdrop: "static",
            keyboard: false
        });
    }
};

// Initial
WorkspaceHome.init();
WorkspaceModal.init();
ReportModal.init();
ReportProgressModal.init();
AuthModal.init();

// Update report progress
eel.expose(updateReportProgress);
function updateReportProgress(msg, percent) {
    a5.body.append(msg + "<br>");
    ReportProgressModal.updateProgressBar(percent);
}

eel.expose(askLoginInfo);
function askLoginInfo() {
    AuthModal.invoke();
}

// Onload
window.onload = async function(e) {
    await translate();

    var data = await eel.get_workspace_data()();
    WorkspaceHome.render(data);

    $(".preloader").fadeOut();

    WorkspaceHome.updateFunctionInfo();
};
