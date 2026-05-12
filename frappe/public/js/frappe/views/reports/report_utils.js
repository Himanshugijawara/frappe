frappe.provide("frappe.report_utils");

frappe.report_utils = {
	make_chart_options: function (
		columns,
		raw_data,
		{ y_fields, x_field, chart_type, colors, height }
	) {
		const type = chart_type.toLowerCase();

		let rows = raw_data.result.filter((value) => Object.keys(value).length);

		let labels = get_column_values(x_field);
		let datasets = y_fields.map((y_field) => ({
			name: get_translated_column_label(y_field),
			values: get_column_values(y_field).map((d) => Number(d)),
		}));

		if (raw_data.add_total_row) {
			labels = labels.slice(0, -1);
			datasets.forEach((dataset) => {
				dataset.values = dataset.values.slice(0, -1);
			});
		}

		return {
			data: {
				labels: labels.length ? labels : null,
				datasets: datasets,
			},
			truncateLegends: 1,
			type: type,
			height: height ? height : 280,
			colors: colors,
			axisOptions: {
				shortenYAxisNumbers: 1,
				xAxisMode: "tick",
				numberFormatter: frappe.utils.format_chart_axis_number,
			},
		};

		function get_column_values(column_name) {
			if (Array.isArray(rows[0])) {
				let column_index = columns.findIndex((column) => column.fieldname == column_name);
				return rows.map((row) => row[column_index]);
			} else {
				return rows.map((row) => row[column_name]);
			}
		}

		function get_translated_column_label(fieldname) {
			let column = columns.find((column) => column.fieldname === fieldname);
			return column?.label ?? __(frappe.model.unscrub(fieldname));
		}
	},

	get_field_options_from_report: function (columns, data) {
		const rows = data.result.filter((value) => Object.keys(value).length);
		const first_row = Array.isArray(rows[0])
			? rows[0]
			: columns.map((col) => rows[0][col.fieldname]);

		const indices = first_row.reduce((accumulator, current_value, current_index) => {
			if (Number.isFinite(current_value)) {
				accumulator.push(current_index);
			}
			return accumulator;
		}, []);

		function get_options(fields) {
			return fields.map((field) => {
				if (field.fieldname) {
					return { label: field.label, value: field.fieldname };
				} else {
					field = frappe.report_utils.prepare_field_from_column(field);
					return { label: field.label, value: field.fieldname };
				}
			});
		}

		const numeric_fields = columns.filter((col, i) => indices.includes(i));
		const non_numeric_fields = columns.filter((col, i) => !indices.includes(i));

		let numeric_field_options = get_options(numeric_fields);
		let non_numeric_field_options = get_options(non_numeric_fields);

		return {
			numeric_fields: numeric_field_options,
			non_numeric_fields: non_numeric_field_options,
		};
	},

	prepare_field_from_column: function (column) {
		if (typeof column === "string") {
			if (column.includes(":")) {
				let [label, fieldtype, width] = column.split(":");
				let options;

				if (fieldtype.includes("/")) {
					[fieldtype, options] = fieldtype.split("/");
				}

				column = {
					label,
					fieldname: label,
					fieldtype,
					width,
					options,
				};
			} else {
				column = {
					label: column,
					fieldname: column,
					fieldtype: "Data",
				};
			}
		}
		return column;
	},

	get_report_filters: function (report_name) {
		if (frappe.query_reports[report_name]) {
			let filters = frappe.query_reports[report_name].filters;
			return Promise.resolve(filters);
		}

		return frappe
			.xcall("frappe.desk.query_report.get_script", {
				report_name: report_name,
			})
			.then((r) => {
				frappe.dom.eval(r.script);
				return frappe.after_ajax(() => {
					if (
						frappe.query_reports[report_name] &&
						!frappe.query_reports[report_name].filters &&
						r.filters
					) {
						return (frappe.query_reports[report_name].filters = r.filters);
					}
					return (
						frappe.query_reports[report_name] &&
						frappe.query_reports[report_name].filters
					);
				});
			});
	},

	get_filter_values(filters) {
		return filters
			.map((f) => {
				var v = f.default;
				return {
					[f.fieldname]: v,
				};
			})
			.reduce((acc, f) => {
				Object.assign(acc, f);
				return acc;
			}, {});
	},

	get_result_of_fn(fn, values) {
		const get_result = {
			Minimum: (values) => values.reduce((min, val) => Math.min(min, val), values[0]),
			Maximum: (values) => values.reduce((min, val) => Math.max(min, val), values[0]),
			Average: (values) => values.reduce((a, b) => a + b, 0) / values.length,
			Sum: (values) => values.reduce((a, b) => a + b, 0),
		};
		return get_result[fn](values);
	},

	get_export_dialog(report_name, extra_fields, callback) {
		const fields = [
			{
				label: "File Format",
				fieldname: "file_format",
				fieldtype: "Select",
				options: ["Excel", "CSV"],
				default: "Excel",
				reqd: 1,
			},
			{
				label: __("Export in Background"),
				fieldname: "export_in_background",
				fieldtype: "Check",
			},
			// ─── Excel Styling (everything below collapses together) ────────
			{
				fieldtype: "Section Break",
				fieldname: "xlsx_style_section",
				label: __("Excel Styling"),
				collapsible: 1,
				collapsible_depends_on: "eval:1",
				depends_on: "eval:doc.file_format=='Excel'",
			},
			// Header subsection
			{
				fieldtype: "HTML",
				fieldname: "xlsx_header_heading",
				options: `<div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;margin-top:4px;">${__("Header")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_style_header_enable",
				label: __("Customize header"),
				default: 0,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_header_bg_color",
				label: __("Header background"),
				default: "#4472C4",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_header_enable",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_header_font_color",
				label: __("Header text color"),
				default: "#FFFFFF",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_header_enable",
			},
			{
				fieldtype: "Int",
				fieldname: "xlsx_header_font_size",
				label: __("Header font size"),
				default: 11,
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_header_enable",
			},
			// Data cells subsection
			{
				fieldtype: "HTML",
				fieldname: "xlsx_data_heading",
				options: `<hr style="margin:12px 0 4px 0;"><div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;">${__("Data Cells")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_style_data_enable",
				label: __("Customize data cells"),
				default: 0,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_data_bg_color",
				label: __("Data background"),
				default: "#FFFFFF",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_data_enable",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_data_font_color",
				label: __("Data text color"),
				default: "#222222",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_data_enable",
			},
			{
				fieldtype: "Int",
				fieldname: "xlsx_data_font_size",
				label: __("Data font size"),
				default: 11,
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_style_data_enable",
			},
			// Borders subsection
			{
				fieldtype: "HTML",
				fieldname: "xlsx_borders_heading",
				options: `<hr style="margin:12px 0 4px 0;"><div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;">${__("Borders")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_borders_enable",
				label: __("Add borders"),
				default: 0,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Select",
				fieldname: "xlsx_border_style",
				label: __("Border style"),
				options: ["thin", "medium", "thick", "dashed", "dotted", "double", "hair"],
				default: "thin",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_borders_enable",
			},
			{
				fieldtype: "Select",
				fieldname: "xlsx_border_scope",
				label: __("Apply borders to"),
				options: [
					{ value: "all", label: __("Header + Data") },
					{ value: "header_only", label: __("Header only") },
					{ value: "data_only", label: __("Data only") },
				],
				default: "all",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_borders_enable",
			},
			// Zebra subsection
			{
				fieldtype: "HTML",
				fieldname: "xlsx_zebra_heading",
				options: `<hr style="margin:12px 0 4px 0;"><div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;">${__("Alternate Row Colors")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_zebra_enable",
				label: __("Alternate row colors"),
				default: 0,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_zebra_color",
				label: __("Stripe background"),
				default: "#F2F2F2",
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_zebra_enable",
			},
			{
				fieldtype: "Color",
				fieldname: "xlsx_zebra_font_color",
				label: __("Stripe text color"),
				description: __("Optional. Leave blank to keep the data text color."),
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_zebra_enable",
			},
			// Number formatting subsection
			{
				fieldtype: "HTML",
				fieldname: "xlsx_numfmt_heading",
				options: `<hr style="margin:12px 0 4px 0;"><div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;">${__("Number Formatting")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_numfmt_enable",
				label: __("Customize numeric columns"),
				description: __(
					"Applies to Currency, Float, Percent and Int columns"
				),
				default: 0,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Int",
				fieldname: "xlsx_numfmt_precision",
				label: __("Decimal places"),
				default: 2,
				description: __("0–20. Integer columns always render as whole numbers."),
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_numfmt_enable",
			},
			{
				fieldtype: "Check",
				fieldname: "xlsx_numfmt_right_align",
				label: __("Right-align numeric columns"),
				default: 1,
				depends_on: "eval:doc.file_format=='Excel' && doc.xlsx_numfmt_enable",
			},
			// Live preview
			{
				fieldtype: "HTML",
				fieldname: "xlsx_preview_heading",
				options: `<hr style="margin:12px 0 4px 0;"><div class="text-muted small font-weight-bold" style="text-transform:uppercase;letter-spacing:0.5px;">${__("Preview")}</div>`,
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "HTML",
				fieldname: "xlsx_preview",
				depends_on: "eval:doc.file_format=='Excel'",
			},
			{
				fieldtype: "Section Break",
				fieldname: "csv_settings",
				label: "Settings",
				collapsible: 1,
				depends_on: "eval:doc.file_format=='CSV'",
			},
			{
				fieldtype: "Data",
				label: "CSV Delimiter",
				fieldname: "csv_delimiter",
				default: ",",
				length: 1,
				depends_on: "eval:doc.file_format=='CSV'",
			},
			{
				fieldtype: "Select",
				label: "CSV Quoting",
				fieldname: "csv_quoting",
				options: [
					{ value: 0, label: "Minimal" },
					{ value: 1, label: "All" },
					{ value: 2, label: "Non-numeric" },
					{ value: 3, label: "None" },
				],
				default: 2,
				depends_on: "eval:doc.file_format=='CSV'",
			},
			{
				fieldtype: "Data",
				label: "CSV Decimal Separator",
				fieldname: "csv_decimal_sep",
				default: ".",
				length: 1,
				depends_on: "eval:doc.file_format=='CSV' && doc.csv_quoting != 2",
			},
			{
				fieldtype: "Small Text",
				label: "CSV Preview",
				fieldname: "csv_preview",
				read_only: 1,
				depends_on: "eval:doc.file_format=='CSV'",
			},
		];

		if (extra_fields) {
			fields.push(
				{
					fieldtype: "Section Break",
					fieldname: "extra_fields",
					collapsible: 0,
				},
				...extra_fields
			);
		}

		const dialog = new frappe.ui.Dialog({
			title: __("Export Report: {0}", [report_name], "Export report"),
			fields: fields,
			primary_action_label: __("Download", null, "Export report"),
			primary_action: callback,
		});

		function update_csv_preview(dialog) {
			const is_query_report = frappe.get_route()[0] === "query-report";
			const report = is_query_report ? frappe.query_report : cur_list;
			const columns = report.columns.filter((col) => col.hidden !== 1);
			let PREVIEW_DATA = [
				columns.map((col) => __(is_query_report ? col.label : col.name)),
				...report.data
					.slice(0, 3)
					.map((row) =>
						columns.map((col) => row[is_query_report ? col.fieldname : col.field])
					),
			];

			dialog.set_value(
				"csv_preview",
				frappe.report_utils.get_csv_preview(
					PREVIEW_DATA,
					dialog.get_value("csv_quoting"),
					dialog.get_value("csv_delimiter"),
					dialog.get_value("csv_decimal_sep")
				)
			);
		}

		function update_xlsx_preview(dialog) {
			const wrapper = dialog.fields_dict["xlsx_preview"]?.$wrapper;
			if (!wrapper) return;

			const v = dialog.get_values(true) || {};

			const header_enabled = !!v.xlsx_style_header_enable;
			const data_enabled = !!v.xlsx_style_data_enable;
			const borders_enabled = !!v.xlsx_borders_enable;
			const zebra_enabled = !!v.xlsx_zebra_enable;
			const numfmt_enabled = !!v.xlsx_numfmt_enable;

			// header style
			const header_bg = header_enabled ? v.xlsx_header_bg_color || "#4472C4" : "transparent";
			const header_fg = header_enabled ? v.xlsx_header_font_color || "#FFFFFF" : "inherit";
			const header_size = header_enabled ? cint(v.xlsx_header_font_size) || 11 : 11;

			// data style
			const data_bg = data_enabled ? v.xlsx_data_bg_color || "#FFFFFF" : "transparent";
			const data_fg = data_enabled ? v.xlsx_data_font_color || "#222222" : "inherit";
			const data_size = data_enabled ? cint(v.xlsx_data_font_size) || 11 : 11;

			// borders
			const border_scope = v.xlsx_border_scope || "all";
			const header_has_border =
				borders_enabled && (border_scope === "all" || border_scope === "header_only");
			const data_has_border =
				borders_enabled && (border_scope === "all" || border_scope === "data_only");
			const border_css = "1px solid #6c757d";
			const header_border_css = header_has_border ? border_css : "none";
			const data_border_css = data_has_border ? border_css : "none";

			// zebra
			const zebra_color = zebra_enabled ? v.xlsx_zebra_color || "#F2F2F2" : null;
			const zebra_font = zebra_enabled ? v.xlsx_zebra_font_color || null : null;

			// number formatting
			const right_align_numbers = numfmt_enabled && !!v.xlsx_numfmt_right_align;
			const numeric_align = right_align_numbers ? "right" : "left";
			let precision = 2;
			if (numfmt_enabled) {
				const p = cint(v.xlsx_numfmt_precision);
				if (p >= 0 && p <= 20) precision = p;
			}
			const fmt_float = (n) => (numfmt_enabled ? n.toFixed(precision) : n.toString());

			const cell_pad = "padding: 6px 10px;";
			const header_style = `background:${header_bg};color:${header_fg};font-weight:600;font-size:${header_size}px;border:${header_border_css};${cell_pad}`;
			const data_cell_style = (striped, align = "left") => {
				const bg = striped && zebra_color ? zebra_color : data_bg;
				const fg = striped && zebra_font ? zebra_font : data_fg;
				return `background:${bg};color:${fg};font-size:${data_size}px;border:${data_border_css};text-align:${align};${cell_pad}`;
			};

			// header alignment matches data alignment for numeric columns (mirrors XLSXStyleBuilder)
			const header_style_aligned = (align = "left") =>
				header_style + `text-align:${align};`;

			wrapper.html(`
				<table style="border-collapse:collapse;width:100%;font-size:13px;">
					<thead>
						<tr>
							<th style="${header_style_aligned("left")}">${__("Name")}</th>
							<th style="${header_style_aligned(numeric_align)}">${__("Score")}</th>
							<th style="${header_style_aligned("left")}">${__("Status")}</th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<td style="${data_cell_style(false, "left")}">Alpha</td>
							<td style="${data_cell_style(false, numeric_align)}">${fmt_float(100)}</td>
							<td style="${data_cell_style(false, "left")}">Open</td>
						</tr>
						<tr>
							<td style="${data_cell_style(true, "left")}">Beta</td>
							<td style="${data_cell_style(true, numeric_align)}">${fmt_float(200.5)}</td>
							<td style="${data_cell_style(true, "left")}">Closed</td>
						</tr>
						<tr>
							<td style="${data_cell_style(false, "left")}">Gamma</td>
							<td style="${data_cell_style(false, numeric_align)}">${fmt_float(300.123)}</td>
							<td style="${data_cell_style(false, "left")}">Open</td>
						</tr>
					</tbody>
				</table>
			`);
		}

		dialog.fields_dict["file_format"].df.onchange = () => {
			update_csv_preview(dialog);
			update_xlsx_preview(dialog);
		};
		dialog.fields_dict["csv_quoting"].df.onchange = () => update_csv_preview(dialog);
		dialog.fields_dict["csv_delimiter"].df.onchange = () => {
			if (!dialog.get_value("csv_delimiter")) {
				dialog.set_value("csv_delimiter", ",");
			}
			update_csv_preview(dialog);
		};
		dialog.fields_dict["csv_decimal_sep"].df.onchange = () => {
			if (!dialog.get_value("csv_decimal_sep")) {
				dialog.set_value("csv_decimal_sep", ".");
			}
			update_csv_preview(dialog);
		};

		// hook every XLSX styling field to refresh the preview
		const xlsx_preview_fields = [
			"xlsx_style_header_enable",
			"xlsx_header_bg_color",
			"xlsx_header_font_color",
			"xlsx_header_font_size",
			"xlsx_style_data_enable",
			"xlsx_data_bg_color",
			"xlsx_data_font_color",
			"xlsx_data_font_size",
			"xlsx_borders_enable",
			"xlsx_border_style",
			"xlsx_border_scope",
			"xlsx_zebra_enable",
			"xlsx_zebra_color",
			"xlsx_zebra_font_color",
			"xlsx_numfmt_enable",
			"xlsx_numfmt_precision",
			"xlsx_numfmt_right_align",
		];
		xlsx_preview_fields.forEach((fieldname) => {
			const field = dialog.fields_dict[fieldname];
			if (field) {
				field.df.onchange = () => update_xlsx_preview(dialog);
			}
		});

		// render the initial preview once the dialog is shown
		const _orig_show = dialog.show.bind(dialog);
		dialog.show = function () {
			_orig_show();
			update_xlsx_preview(dialog);
		};

		return dialog;
	},

	build_xlsx_user_style(values) {
		/**
		 * Convert flat dialog field values into the structured dict expected by
		 * `frappe.utils.xlsxutils.apply_user_styles` on the backend.
		 *
		 * Returns null when no styling has been opted-in, so the caller can omit
		 * the form param entirely and preserve existing default-export behaviour.
		 */
		if (!values || values.file_format !== "Excel") return null;

		const style = {};

		if (values.xlsx_style_header_enable) {
			const header = {};
			if (values.xlsx_header_bg_color) header.bg_color = values.xlsx_header_bg_color;
			if (values.xlsx_header_font_color) header.font_color = values.xlsx_header_font_color;
			const font_size = cint(values.xlsx_header_font_size);
			if (font_size > 0) header.font_size = font_size;
			if (Object.keys(header).length) style.header = header;
		}

		if (values.xlsx_style_data_enable) {
			const data = {};
			if (values.xlsx_data_bg_color) data.bg_color = values.xlsx_data_bg_color;
			if (values.xlsx_data_font_color) data.font_color = values.xlsx_data_font_color;
			const font_size = cint(values.xlsx_data_font_size);
			if (font_size > 0) data.font_size = font_size;
			if (Object.keys(data).length) style.data = data;
		}

		if (values.xlsx_borders_enable) {
			style.borders = {
				style: values.xlsx_border_style || "thin",
				scope: values.xlsx_border_scope || "all",
			};
		}

		if (values.xlsx_zebra_enable && values.xlsx_zebra_color) {
			const zebra = { color: values.xlsx_zebra_color };
			if (values.xlsx_zebra_font_color) zebra.font_color = values.xlsx_zebra_font_color;
			style.zebra_stripes = zebra;
		}

		if (values.xlsx_numfmt_enable) {
			const numfmt = {};
			const precision = cint(values.xlsx_numfmt_precision);
			if (precision >= 0 && precision <= 20) numfmt.precision = precision;
			if (values.xlsx_numfmt_right_align) numfmt.right_align_numeric = true;
			if (Object.keys(numfmt).length) style.number_formatting = numfmt;
		}

		return Object.keys(style).length ? style : null;
	},

	get_csv_preview(data, quoting, delimiter, decimal_sep) {
		// data: array of arrays
		// quoting: 0 - minimal, 1 - all, 2 - non-numeric, 3 - none
		// delimiter: any single character
		quoting = cint(quoting);
		const QUOTING = {
			Minimal: 0,
			All: 1,
			NonNumeric: 2,
			None: 3,
		};

		if (delimiter.length > 1) {
			frappe.throw(__("Delimiter must be a single character"));
		}

		if (decimal_sep.length > 1) {
			frappe.throw(__("Decimal Separator must be a single character"));
		}

		if (0 > quoting || quoting > 3) {
			frappe.throw(__("Quoting must be between 0 and 3"));
		}

		if (decimal_sep !== "." && quoting === QUOTING.NonNumeric) {
			frappe.throw(__("Decimal Separator must be '.' when Quoting is set to Non-numeric"));
		}

		return data
			.map((row) => {
				return row
					.map((col) => {
						if (col === null) {
							return "";
						}

						if (typeof col == "string" && col.includes('"')) {
							col = col.replace(/"/g, '""');
						}

						if (typeof col == "number" && decimal_sep !== ".") {
							col = col.toString().replace(".", decimal_sep);
						}

						switch (quoting) {
							case QUOTING.Minimal:
								return typeof col === "string" && col.includes(delimiter)
									? `"${col}"`
									: `${col}`;
							case QUOTING.All:
								return `"${col}"`;
							case QUOTING.NonNumeric:
								return isNaN(col) ? `"${col}"` : `${col}`;
							case QUOTING.None:
								return `${col}`;
						}
					})
					.join(delimiter);
			})
			.join("\n");
	},
};
