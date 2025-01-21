// Copyright (c) 2016, Frappe Technologies Pvt. Ltd. and contributors
// For license information, please see license.txt

frappe.ui.form.on("Department", {
	onload: function (frm) {
		frm.set_query("parent_department", function () {
			return { filters: [["Department", "is_group", "=", 1]] };
		});
	},
	refresh: function (frm) {
		frappe.call({
			method: "frappe.client.get_list",
			args: {
				"doctype": "Employee",
				"fields": ["name","user_id", "department", "company" ],
				"filters": [
					["department", "=", frm.doc.name],
					["company", "=", frm.doc.company]
				]
			},
			callback: function(r) {
				table = "";
				if (r.message && r.message.length > 0) {
					// Tạo bảng HTML với class
					var rows = r.message;
					table += "<table class='table table-bordered' style='width:100%; border-collapse: collapse;'>";
					table += "<thead><tr style='background-color: #f2f2f2;'>";

					// Thêm tiêu đề cột với class
					Object.keys(rows[0]).forEach(function(col) {
						table += `<th class='table-header' style='padding: 8px; text-align: left;'>${col}</th>`;
					});
					table += "<th class='table-header' style='padding: 8px; text-align: center;'>Actions</th>";
					table += "</tr></thead><tbody>";

					// Thêm dữ liệu vào bảng với class
					rows.forEach(function(row) {
						table += "<tr>";
						Object.keys(row).forEach(function(col) {
							table += `<td class='table-cell' style='padding: 8px;'>${row[col]}</td>`;
						});
						table += `<td class='table-cell' style='padding: 8px; text-align: center;'>
                            <button class='btn btn-primary edit-btn' data-id='${row.name}'>
                                <i class='fa fa-edit'></i> Edit
                            </button>
                          </td>`;
						table += "</tr>";
					});

					table += "</tbody></table>";
				}
				table += "<tfoot><tr><td colspan='4' style='text-align: center;'><button class='btn btn-primary' id='load_data_btn'>Add Employee</button></td></tr></tfoot>";
				// Cập nhật nội dung của trường HTML
				frm.set_df_property("custom_table", "options", table);
				frm.refresh_field("custom_table");

				$("#load_data_btn").click(function() {
					if(frm.is_dirty()){
						frappe.msgprint(__("Please save the form first."))
					}else{
						frappe.new_doc("Employee", {
							department: frm.doc.name,
							company: frm.doc.company
						}).then(doc => {
							frappe.set_route("Form", "Employee", doc.name);
						})
					}	
				})

				$(".edit-btn").click(function() {
					frappe.set_route("Form", "Employee", $(this).data("id"));
				})
			}
		});

		frm.add_custom_button(__("Go to sync m365"), function()  {
			if(frm.is_dirty()){
				frappe.msgprint(__("Please save the form first."))
			}else{
				let mailnickname = frm.doc.department_name.toLowerCase().replace(/\s+/g, "");
					frappe.new_doc("M365 Groups", {
						"m365_group_name": frm.doc.department_name,
						"mailnickname": mailnickname,
						"department_id": frm.doc.name,
						"company": frm.doc.company
					}).then( doc => {
						frappe.set_route("Form", "M365 Groups", doc.name);												
					});
					window.open(r.message, "_blank");
			}
			});

		if (!frm.doc.parent_department && !frm.is_new()) {
			frm.set_read_only();
			frm.set_intro(__("This is a root department and cannot be edited."));
		}
	},
	validate: function (frm) {
		if (frm.doc.name == "All Departments") {
			frappe.throw(__("You cannot edit root node."));
		}
	},
});