import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox
import openpyxl
from tkinter import filedialog
from ttkbootstrap import Style
from ttkbootstrap.widgets import DateEntry, Combobox
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Google Sheets Setup ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("KSClothing").sheet1

# --- Add Order UI ---
def add_order_ui():
    add_win = ttk.Toplevel()
    add_win.title("Add Order")
    add_win.geometry("420x620")
    add_win.resizable(False, False)

    all_orders = sheet.col_values(2)
    existing_tracking_ids = sheet.col_values(3)

    if len(all_orders) > 1:
        last_order_id = all_orders[-1]
        last_num = int(last_order_id[2:])
        new_num = last_num + 1
    else:
        new_num = 1

    order_id = f"KS{new_num:02d}"
    date_str = datetime.date.today().strftime('%Y-%m-%d')

    ttk.Label(add_win, text=f"📅 Date: {date_str}", font=('Segoe UI', 10)).pack(pady=(10, 5))
    ttk.Label(add_win, text=f"🆔 Order ID: {order_id}", font=('Segoe UI', 11, "bold"), bootstyle="info").pack(pady=5)

    def focus_next(entry):
        return lambda event: entry.focus_set()

    labels_entries = [
        ("🚚 Tracking ID", ""),
        ("👤 Customer Name", ""),
        ("📞 Phone Number", ""),
        ("📦 Product Details", ""),
        ("🔢 Quantity", ""),
        ("💵 Delivery Cost (LKR)", ""),
        ("💰 Total Price (LKR)", "")
        # Removed COD from here
    ]

    entries = []
    for label_text, _ in labels_entries:
        ttk.Label(add_win, text=label_text, font=('Segoe UI', 9)).pack(pady=(8, 0))
        entry = ttk.Entry(add_win, font=('Segoe UI', 10))
        entry.pack(padx=20, fill=X)
        entries.append(entry)

    # COD Payment dropdown
    ttk.Label(add_win, text="💳 COD Payment", font=('Segoe UI', 9)).pack(pady=(8, 0))
    cod_combo = ttk.Combobox(add_win, values=["Yes", "No"], state="readonly", font=('Segoe UI', 10))
    cod_combo.set("No")
    cod_combo.pack(padx=20, fill=X)

    for i in range(len(entries) - 1):
        entries[i].bind("<Return>", focus_next(entries[i + 1]))
    entries[-1].bind("<Return>", lambda e: cod_combo.focus_set())
    cod_combo.bind("<Return>", lambda e: submit_order())

    def submit_order():
        city_pack_id = entries[0].get().strip()
        if not city_pack_id:
            messagebox.showerror("Error", "Tracking ID cannot be empty.")
            return
        if city_pack_id in existing_tracking_ids:
            messagebox.showerror("Error", "Tracking ID already exists.")
            return

        try:
            phone = str(int(entries[2].get().strip()))
            qty = str(int(entries[4].get().strip()))
            delivery_cost = str(int(entries[5].get().strip()))
            total_price = str(int(entries[6].get().strip()))
        except ValueError:
            messagebox.showerror("Error", "Phone, Quantity, Delivery, and Total must be numbers.")
            return

        cod_input = cod_combo.get().strip().capitalize()
        if cod_input not in ["Yes", "No"]:
            messagebox.showerror("Error", "COD Payment must be 'Yes' or 'No'.")
            return

        cod_bool = "TRUE" if cod_input == "Yes" else "FALSE"

        row = [
            date_str,
            order_id,
            city_pack_id,
            entries[1].get().strip(),
            phone,
            entries[3].get().strip(),
            qty,
            delivery_cost,
            total_price,
            cod_bool
        ]

        sheet.append_row(row)
        messagebox.showinfo("Success", f"✅ Order {order_id} added.")
        add_win.destroy()

    ttk.Button(add_win, text="✅ Submit Order", command=submit_order, bootstyle=SUCCESS, width=25).pack(pady=20)


def update_order_ui():
    update_win = ttk.Toplevel()
    update_win.title("✏️ Update Order")
    update_win.geometry("420x700")

    # --- SEARCH SECTION ---
    search_frame = ttk.Frame(update_win)
    search_frame.pack(pady=10)

    ttk.Label(search_frame, text="Enter Order ID to Update").pack(pady=5)
    entry_order_id = ttk.Entry(search_frame, width=30)
    entry_order_id.pack(pady=5)

    # --- SCROLLABLE FRAME SETUP ---
    canvas = ttk.Canvas(update_win, borderwidth=0, highlightthickness=0, height=580)
    scrollbar = ttk.Scrollbar(update_win, orient="vertical", command=canvas.yview)
    content_frame = ttk.Frame(canvas)

    content_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=content_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def fetch_order():
        order_id = entry_order_id.get().strip()
        all_orders = sheet.get_all_values()
        for idx, row in enumerate(all_orders):
            if row[1] == order_id:
                fill_fields(row, idx + 1)
                return
        messagebox.showerror("Not Found", f"No order found with ID: {order_id}")

    def fill_fields(fields, row_index):
        for widget in content_frame.winfo_children():
            widget.destroy()

        ttk.Label(content_frame, text="Edit the fields below").pack(pady=10)

        entry_widgets = []
        labels = [
            "Date", "Order ID", "Tracking ID", "Customer Name", "Phone",
            "Product", "Quantity", "Delivery Cost", "Total Price", "COD Payment"
        ]

        for i, label in enumerate(labels):
            ttk.Label(content_frame, text=label).pack(anchor='w', padx=20)

            if i == 9:  # COD Payment
                cod_combo = ttk.Combobox(content_frame, values=["Yes", "No"], state="readonly")
                cod_combo.set("Yes" if fields[i].strip().lower() == "true" else "No")
                cod_combo.pack(padx=20, pady=3, fill='x')
                entry_widgets.append(cod_combo)
            else:
                entry = ttk.Entry(content_frame)
                entry.insert(0, fields[i])
                entry.pack(padx=20, pady=3, fill='x')

                if i in (0, 1):  # Disable Date and Order ID
                    entry.config(state='disabled')

                entry_widgets.append(entry)

        def save_changes():
            updated_row = []
            for i, widget in enumerate(entry_widgets):
                if i in (0, 1):
                    updated_row.append(fields[i])
                elif i == 9:
                    updated_row.append("TRUE" if widget.get() == "Yes" else "FALSE")
                else:
                    updated_row.append(widget.get().strip())

            try:
                sheet.update(f"A{row_index}:J{row_index}", [updated_row])
                messagebox.showinfo("Success", "✅ Order updated successfully.")
                update_win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"❌ Failed to update order:\n{e}")

        ttk.Button(content_frame, text="💾 Save Changes", command=save_changes).pack(pady=20)

    # 🔍 Search button placed AFTER fetch_order is defined
    ttk.Button(search_frame, text="🔍 Fetch Order", command=fetch_order).pack(pady=10)



def search_order_ui():
    search_win = ttk.Toplevel()
    search_win.title("🔍 Search Order")
    search_win.geometry("500x400")

    ttk.Label(search_win, text="Enter Order ID or Tracking ID:", font=('Helvetica', 10)).pack(pady=10)
    search_entry = ttk.Entry(search_win, width=40)
    search_entry.pack()

    result_text = ttk.Text(search_win, width=60, height=15, state='disabled', font=('Courier', 10))
    result_text.pack(pady=10)

    def perform_search():
        query = search_entry.get().strip()
        if not query:
            messagebox.showerror("Input Error", "Please enter an Order ID or Tracking ID.")
            return

        data = sheet.get_all_values()
        headers = data[0]
        found = False

        for row in data[1:]:
            if query == row[1] or query == row[2]:  # Match Order ID or Tracking ID
                found = True
                result_text.config(state='normal')
                result_text.delete("1.0", ttk.END)
                for i, value in enumerate(row):
                    result_text.insert(ttk.END, f"{headers[i]}: {value}\n")
                result_text.config(state='disabled')
                break

        if not found:
            result_text.config(state='normal')
            result_text.delete("1.0", ttk.END)
            result_text.insert(ttk.END, "❌ No order found with that ID.")
            result_text.config(state='disabled')

    ttk.Button(search_win, text="Search", command=perform_search).pack(pady=5)


def view_orders_ui():
    view_win = ttk.Toplevel()
    view_win.title("📄 View Orders")
    view_win.geometry("900x600")

    filtered_rows = []  # <-- Store filtered data here

    # Filter section
    filter_frame = ttk.Frame(view_win)
    filter_frame.pack(pady=10, fill=ttk.X)

    ttk.Label(filter_frame, text="COD Payment:").grid(row=0, column=0, padx=5)
    cod_var = ttk.StringVar(value="All")
    cod_dropdown = ttk.OptionMenu(filter_frame, cod_var, "All", "Yes", "No")
    cod_dropdown.grid(row=0, column=1, padx=5)

    ttk.Label(filter_frame, text="From Date (YYYY-MM-DD):").grid(row=0, column=2, padx=5)
    from_date_entry = ttk.Entry(filter_frame)
    from_date_entry.grid(row=0, column=3, padx=5)

    ttk.Label(filter_frame, text="To Date (YYYY-MM-DD):").grid(row=0, column=4, padx=5)
    to_date_entry = ttk.Entry(filter_frame)
    to_date_entry.grid(row=0, column=5, padx=5)

    frame = ttk.Frame(view_win)
    frame.pack(fill=ttk.BOTH, expand=True, padx=10, pady=10)

    result_text = ttk.Text(frame, wrap=ttk.NONE, font=('Courier', 10))
    result_text.pack(side=ttk.LEFT, fill=ttk.BOTH, expand=True)

    yscroll = ttk.Scrollbar(frame, orient='vertical', command=result_text.yview)
    yscroll.pack(side=ttk.RIGHT, fill=ttk.Y)
    result_text.config(yscrollcommand=yscroll.set)

    def filter_orders():
        nonlocal filtered_rows
        result_text.delete("1.0", ttk.END)
        data = sheet.get_all_values()
        headers = data[0]
        filtered_rows = []  # reset filtered results

        cod_filter = cod_var.get()
        from_date = from_date_entry.get().strip()
        to_date = to_date_entry.get().strip()

        try:
            from_date_obj = datetime.datetime.strptime(from_date, '%Y-%m-%d') if from_date else None
            to_date_obj = datetime.datetime.strptime(to_date, '%Y-%m-%d') if to_date else None
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter dates in YYYY-MM-DD format.")
            return

        for row in data[1:]:
            try:
                row_date = datetime.datetime.strptime(row[0].strip(), '%Y-%m-%d')
            except:
                continue

            cod_val_raw = row[9].strip().lower()
            cod_val = "yes" if cod_val_raw in ["yes", "true", "1"] else "no"

            cod_match = (
                cod_filter == "All" or
                (cod_filter == "Yes" and cod_val == "yes") or
                (cod_filter == "No" and cod_val == "no")
            )
            date_match = (
                (not from_date_obj or row_date >= from_date_obj) and
                (not to_date_obj or row_date <= to_date_obj)
            )

            if cod_match and date_match:
                filtered_rows.append(row)
                order_data = ", ".join([f"{headers[i]}: {val}" for i, val in enumerate(row)])
                result_text.insert(ttk.END, f"{order_data}\n{'-'*100}\n")

        if not filtered_rows:
            result_text.insert(ttk.END, "⚠️ No orders found for selected filters.\n")

    def export_to_excel():
        if not filtered_rows:
            messagebox.showwarning("No Data", "No filtered data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")],
                                                 title="Save filtered orders")
        if not file_path:
            return  # Cancelled

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            data = sheet.get_all_values()
            headers = data[0]

            ws.append(headers)
            for row in filtered_rows:
                ws.append(row)

            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"Filtered data exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred:\n{e}")

    # Buttons
    button_frame = ttk.Frame(view_win)
    button_frame.pack(pady=10)

    ttk.Button(button_frame, text="Apply Filter", command=filter_orders).grid(row=0, column=0, padx=5)
    ttk.Button(button_frame, text="Export to Excel", command=export_to_excel).grid(row=0, column=1, padx=5)

    filter_orders()  # Load all on startup



# --- Main App UI ---
def main():
    root = ttk.Window(themename="flatly")
    root.title("KSClothing Order Manager")
    root.geometry("360x400")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=20)
    frame.pack(expand=True)

    ttk.Label(
        frame,
        text="🧾 KS Clothing Order Manager",
        font=("Segoe UI", 16, "bold"),
        bootstyle="dark"
    ).pack(pady=(10, 30))

    button_style = {
        "width": 30,
        "padding": 10
    }

    ttk.Button(frame, text="➕ Add Order", command=add_order_ui, bootstyle=PRIMARY, **button_style).pack(pady=5)
    ttk.Button(frame, text="✏️ Update Order", command=update_order_ui, bootstyle=WARNING, **button_style).pack(pady=5)
    ttk.Button(frame, text="🔍 Search Order by ID", command=search_order_ui, bootstyle=INFO, **button_style).pack(pady=5)
    ttk.Button(frame, text="📄 View All Orders", command=view_orders_ui, bootstyle=SUCCESS, **button_style).pack(pady=5)
    ttk.Button(frame, text="❌ Exit", command=root.destroy, bootstyle=DANGER, **button_style).pack(pady=(20, 0))

    root.mainloop()

if __name__ == "__main__":
    main()
