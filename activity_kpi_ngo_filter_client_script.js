frappe.ui.form.on('Project proposal', {
    refresh(frm) {
        // Filter Activity KPI dropdown by selected NGO partner
        if (frm.fields_dict.custom_activity) {
            frm.fields_dict.custom_activity.get_query = function() {
                return { filters: { ngo: frm.doc.ngo || '' } };
            };
        }
    },
    ngo(frm) {
        // Re-apply filter when NGO partner changes
        if (frm.fields_dict.custom_activity) {
            frm.fields_dict.custom_activity.get_query = function() {
                return { filters: { ngo: frm.doc.ngo || '' } };
            };
            // Clear previously selected activities when partner changes
        }
    }
});
