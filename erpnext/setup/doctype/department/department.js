// Copyright (c) 2016, Frappe Technologies Pvt. Ltd. and contributors
// For license information, please see license.txt

frappe.ui.form.on("Department", {
	onload: function (frm) {
		frm.set_query("parent_department", function () {
			return { filters: [["Department", "is_group", "=", 1]] };
		});
	},
	refresh: function (frm) {
		// read-only for root department

		office_365_logo = `<img src="/assets/frappe/icons/social/office_365.svg" alt="Office 365">`

		frm.add_custom_button(__(`${frm.doc.m365_group ? "Sync ": "Create "}`+ `M365 Group ${office_365_logo}`), function () {
			if (frm.is_dirty()) {
				frappe.msgprint("Please save the form first.")
			} else {

				frappe.call({
					method: "frappe_m365.frappe_m365.doctype.m365_groups.m365_groups.create_m365_group_for_any_doc",
					freeze: 1,
					freeze_message: "<h4>Please wait while we are creating M365 Group...</h4>",
					args:{
						doc: frm.doc,
						members_search_field: "department",
						members_doctype:"Employee"
					},
					callback: function (r) {
						// frm.reload_doc();
					}
				});
			}
		});

		frm.add_custom_button(__("Get Employees"), function () {
			if (frm.is_dirty()) {
				frappe.msgprint("Please save the form first.")
			} else {
				frappe.call({
					method: "get_seperated_members",
					freeze: 1,
					freeze_message: "<h4>Please wait while we get Members on M365 group...</h4>",
					doc: frm.doc,
					callback: function (response) {
						console.log(JSON.stringify(response.message));
						frappe.msgprint(JSON.stringify(response.message));
					}
				});
				
			}
		});

		console.log(frm.doc);

		const iframe = `<iframe src="/app/m365-groups/${frm.doc.m365_group.m365_group_name}" style="width: 100%; height: 600px; border: none;"></iframe>`;
        frm.fields_dict.m365_group_iframe.$wrapper.html(iframe);
		
		frm.trigger("get_seperated_members");
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

	get_seperated_members: function (frm){

		add_erpnext_member_to_m365 = (user_id) => {
			frappe.call({
				method: "add_erpnext_member_to_m365",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				doc: frm.doc,
				args:{
					user_id:user_id
				},
				callback: function (response) {
					console.log(response.message);
					frappe.msgprint(` ${JSON.stringify(response.message)} `);
					frm.reload_doc();
				}
			});
		}

		add_member_to_m365_via_power_automate = (user_id) => {
			frappe.call({
				method: "add_member_to_m365_via_power_automate",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				doc: frm.doc,
				args:{
					user_id:user_id
				},
				callback: function (response) {
					console.log(response.message);
					frappe.msgprint(` ${JSON.stringify(response.message)} `);
					frm.reload_doc();
				}
			});
		}

		unlink_from_department = (employee_name) => {
			frappe.call({
				method: "erpnext.setup.doctype.department.department.unlink_employee_department",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				args:{
					employee_name:employee_name
				},
				callback: function (response) {
					console.log(response.message);
					frappe.msgprint(` ${JSON.stringify(response.message)} `);
					frm.reload_doc();
				}
			});

		}

		add_m365_member_to_erpnext = (email,full_name) => {
			frappe.call({
				method: "create_user_and_employee",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				doc: frm.doc,
				args:{
					email:email,
					full_name:full_name
				},
				callback: function (response) {
					console.log(response.message);
					if(response.message){
						window.location.href = response.message.redirect_to;
					}
					// frappe.msgprint(` ${JSON.stringify(response.message)} `);
					// frm.reload_doc();
				}
			});
		}

		remove_member_from_m365 = (email) => {
			frappe.call({
				method: "remove_member_from_m365",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				doc: frm.doc,
				args:{
					email:email
				},
				callback: function (response) {
					console.log(response.message);
					frappe.msgprint(` ${JSON.stringify(response.message)} `);
					frm.reload_doc();
				}
			});
		}


		frappe.call({
			method: "get_seperated_members",
			freeze: 0,
			freeze_message: "<h4>Please wait while we get Members on M365 group...</h4>",
			doc: frm.doc,
			callback: function (response) {
				console.log(response.message);
				console.log(JSON.stringify(response.message));
				const data = response.message;

				// Tạo bảng HTML
				let erpnext_only_table_html = `
					<table class="table table-bordered">
						<thead>
							<tr>
								<th>Employee ID(Doc Name)</th>
								<th>Name</th>
								<th>Department</th>
								<th>Designation</th>
								<th>Email</th>
								<th>Actions</th>
							</tr>
						</thead>
						<tbody>
				`;

				// Đổ dữ liệu vào bảng
				data.only_in_erpnext.forEach(employee => {
					erpnext_only_table_html += `
						<tr>
							<td><a href="/app/employee/${employee.name}">${employee.name}</td>
							<td>${employee.employee_name}</td>
							<td>${employee.department}</td>
							<td>${employee.designation ?? ""}</td>
							<td>${employee.user_id}</td>
							<td><button onclick = "add_erpnext_member_to_m365('${employee.user_id}')" class = "btn btn-default">Add to M365</button>
							<br><br><button onclick = "unlink_from_department('${employee.name}')" class = "btn btn-default">Unlink from this Department</button>					
						</tr>
					`;
					// <br><br><button onclick = "add_member_to_m365_via_power_automate('${employee.name}')" class = "btn btn-primary">Add to M365 via Power Automate</button></td>
				});

				erpnext_only_table_html += `
						</tbody>
					</table>
				`;

				// Đưa bảng HTML vào trường HTML Field
				frm.fields_dict['erpnext_only_table_html'].$wrapper.html(erpnext_only_table_html);

				let m365_only_members_table_html = `
					<table class="table table-bordered">
						<thead>
							<tr>
								<th>Office 365 ID</th>
								<th>Name</th>
								<th>Designation</th>
								<th>Email</th>
								<th>Actions</th>
								
							</tr>
						</thead>
						<tbody>
				`;

				// Đổ dữ liệu vào bảng
				data.only_in_m365.forEach(member => {
					m365_only_members_table_html += `
						<tr>
							<td>${member.id}</td>
							<td>${member.displayName}</td>
							<td>${member.jobTitle ?? ""}</td>
							<td>${member.mail}</td>
							<td><button onclick = "add_m365_member_to_erpnext('${member.mail}','${member.displayName}')" class = "btn btn-default">Add to this Department</button>
							<br>
							<br><button onclick = "remove_member_from_m365('${member.mail}')" class = "btn btn-default">Remove from M365 Group</button></td>
						</tr>
					`;
				});

				m365_only_members_table_html += `
						</tbody>
					</table>
				`;

				// Đưa bảng HTML vào trường HTML Field
				frm.fields_dict['m365_only_members_table_html'].$wrapper.html(m365_only_members_table_html);

				let both_table_html = `
					<table class="table table-bordered">
						<thead>
							<tr>
								<th>Employee ID(Doc Name)</th>
								<th>Office 365 ID</th>
								<th>Name</th>
								<th>Office Name</th>
								<th>Department</th>
								<th>Designation</th>
								<th>Email</th>
								<th>Actions</th>
							</tr>
						</thead>
						<tbody>
				`;

				// Đổ dữ liệu vào bảng
				data.both.forEach(employee => {
					both_table_html += `
						<tr>
							<td><a href="/app/employee/${employee.name}">${employee.name}</td>
							<td>${employee.office_365_id}</td>
							<td>${employee.employee_name}</td>
							<td>${employee.office_365_name}</td>
							<td>${employee.department}</td>
							<td>${employee.designation ?? ""}</td>
							<td>${employee.user_id}</td>
							<td><button onclick = "unlink_from_department('${employee.name}')" class = "btn btn-default">Unlink from this Department</button>
							<br><br><button onclick = "remove_member_from_m365('${employee.user_id}')" class = "btn btn-default">Remove from M365 Group</button></td>
						</tr>
					`;
				});

				both_table_html += `
						</tbody>
					</table>
				`;

				// Đưa bảng HTML vào trường HTML Field
				frm.fields_dict['both_table_html'].$wrapper.html(both_table_html);
			}
		});
	}
});
