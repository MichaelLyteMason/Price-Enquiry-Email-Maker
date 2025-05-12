import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import csv
import os
import pyperclip

class SteelEnquiryApp:
    def __init__(self, master):
        self.master = master
        master.title("Steel Product Enquiry Manager")
        master.geometry("800x600")

        # Suppliers data file
        self.suppliers_file = 'suppliers.csv'
        self.ensure_suppliers_file()

        # Suppliers list
        self.suppliers = self.load_suppliers()

        # Email template file
        self.email_template_file = 'email_template.txt'
        self.ensure_email_template()

        # Create main UI
        self.create_main_menu()

    def focus_window(self, window):
        """Bring window to front and focus"""
        window.lift()
        window.attributes('-topmost', True)
        window.after_idle(lambda: window.attributes('-topmost', False))
        window.focus_force()

    def create_toplevel(self, title, geometry):
        """Create a properly focused Toplevel window"""
        window = tk.Toplevel(self.master)
        window.title(title)
        window.geometry(geometry)
        window.transient(self.master)  # Associate with main window
        self.focus_window(window)
        return window
    
    def ensure_suppliers_file(self):
        """Ensure suppliers CSV file exists"""
        if not os.path.exists(self.suppliers_file):
            with open(self.suppliers_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Supplier Name', 'Contact Name', 'Email'])

    def ensure_email_template(self):
        """Ensure email template file exists"""
        if not os.path.exists(self.email_template_file):
            with open(self.email_template_file, 'w') as f:
                f.write('''Hi [contact_name],

I hope you're well. I'm emailing to enquire on prices for the following products:
                        

[product_list]

                        
Many thanks,
YOUR NAME''')

    def load_suppliers(self):
        """Load suppliers from CSV, grouped by supplier name"""
        suppliers = {}
        with open(self.suppliers_file, 'r') as f:
            reader = csv.reader(f)
            next(reader)  # Skip header
            for row in reader:
                if row[0] not in suppliers:
                    suppliers[row[0]] = []
                suppliers[row[0]].append(row[1:])
        return suppliers

    def create_main_menu(self):
        """Create the main menu interface"""
        # Clear existing widgets
        for widget in self.master.winfo_children():
            widget.destroy()

        # Main menu frame
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Buttons
        ttk.Button(main_frame, text="Manage Suppliers", 
                   command=self.manage_suppliers_window).grid(row=0, column=0, pady=10)
        
        ttk.Button(main_frame, text="New Enquiry", 
                   command=self.new_enquiry_window).grid(row=1, column=0, pady=10)

    def manage_suppliers_window(self):
        """Open suppliers management window"""
        suppliers_window = self.create_toplevel("Manage Suppliers", "600x800")

        # Create a canvas with scrollbar
        canvas = tk.Canvas(suppliers_window)
        scrollbar = ttk.Scrollbar(suppliers_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Suppliers Table
        columns = ('Supplier Name', 'Contact Name', 'Email')
        tree = ttk.Treeview(scrollable_frame, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=180)

        # Populate table
        def populate_table():
            # Clear existing items
            for i in tree.get_children():
                tree.delete(i)
            
            # Reload suppliers and repopulate
            self.suppliers = self.load_suppliers()
            for supplier, contacts in self.suppliers.items():
                for contact in contacts:
                    tree.insert('', 'end', values=[supplier] + contact)

        populate_table()
        tree.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        # Add Supplier Form
        ttk.Label(scrollable_frame, text="Add New Supplier Contact", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, columnspan=3, pady=(20,10))

        # Supplier Name
        ttk.Label(scrollable_frame, text="Supplier Name:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        supplier_name_entry = ttk.Entry(scrollable_frame, width=40)
        supplier_name_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Contact Name
        ttk.Label(scrollable_frame, text="Contact Name:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        contact_name_entry = ttk.Entry(scrollable_frame, width=40)
        contact_name_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Email
        ttk.Label(scrollable_frame, text="Email:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.E)
        email_entry = ttk.Entry(scrollable_frame, width=40)
        email_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)

        # Add Supplier Button
        def add_supplier():
            supplier_name = supplier_name_entry.get().strip()
            contact_name = contact_name_entry.get().strip()
            email = email_entry.get().strip()

            if not supplier_name or not contact_name or not email:
                messagebox.showerror("Error", "All fields must be filled!")
                return

            # Add to CSV
            with open(self.suppliers_file, 'a', newline='') as f:
                writer = csv.writer(f)
                writer.writerow([supplier_name, contact_name, email])
            
            # Clear entries
            supplier_name_entry.delete(0, tk.END)
            contact_name_entry.delete(0, tk.END)
            email_entry.delete(0, tk.END)

            # Repopulate table
            populate_table()
            
            messagebox.showinfo("Success", "Supplier contact added successfully!")

        ttk.Button(scrollable_frame, text="Add Supplier Contact", command=add_supplier).grid(row=5, column=0, columnspan=3, pady=10)

        # Buttons for existing functionality
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="Edit Supplier Contact", 
                   command=lambda: self.edit_supplier(tree)).grid(row=0, column=0, padx=5)
        
        ttk.Button(button_frame, text="Delete Supplier Contact", 
                   command=lambda: self.delete_supplier(tree)).grid(row=0, column=1, padx=5)

    def edit_supplier(self, tree):
        """Edit selected supplier contact"""
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a supplier contact to edit")
            return

        # Get current values
        current_values = tree.item(selected_item[0])['values']
        
        # Prompt for new details
        supplier_name = simpledialog.askstring("Edit Supplier", "Supplier Name:", initialvalue=current_values[0])
        contact_name = simpledialog.askstring("Edit Supplier", "Contact Name:", initialvalue=current_values[1])
        email = simpledialog.askstring("Edit Supplier", "Email:", initialvalue=current_values[2])

        if supplier_name and contact_name and email:
            # Read entire CSV
            suppliers = []
            with open(self.suppliers_file, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                suppliers.append(header)
                for row in reader:
                    # If this row matches the selected item, replace it
                    if row[0] == current_values[0] and row[1] == current_values[1] and row[2] == current_values[2]:
                        suppliers.append([supplier_name, contact_name, email])
                    else:
                        suppliers.append(row)

            # Write back to CSV
            with open(self.suppliers_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(suppliers)

            # Refresh suppliers list
            self.suppliers = self.load_suppliers()
            
            # Refresh treeview
            for i in tree.get_children():
                tree.delete(i)
            for supplier, contacts in self.suppliers.items():
                for contact in contacts:
                    tree.insert('', 'end', values=[supplier] + contact)

            messagebox.showinfo("Success", "Supplier contact updated successfully!")

    def delete_supplier(self, tree):
        """Delete selected supplier contact"""
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a supplier contact to delete")
            return

        # Get values of selected supplier
        current_values = tree.item(selected_item[0])['values']

        # Confirm deletion
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this supplier contact?"):
            # Read entire CSV
            suppliers = []
            with open(self.suppliers_file, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                suppliers.append(header)
                for row in reader:
                    # Skip the row that matches the selected item
                    if not (row[0] == current_values[0] and 
                            row[1] == current_values[1] and 
                            row[2] == current_values[2]):
                        suppliers.append(row)

            # Write back to CSV
            with open(self.suppliers_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(suppliers)

            # Refresh suppliers list
            self.suppliers = self.load_suppliers()
            
            # Refresh treeview
            for i in tree.get_children():
                tree.delete(i)
            for supplier, contacts in self.suppliers.items():
                for contact in contacts:
                    tree.insert('', 'end', values=[supplier] + contact)

            messagebox.showinfo("Success", "Supplier contact deleted successfully!")

    def new_enquiry_window(self):
        enquiry_window = self.create_toplevel("New Enquiry", "1000x800")

        # === Scrollable canvas setup ===
        canvas = tk.Canvas(enquiry_window)
        scrollbar = ttk.Scrollbar(enquiry_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # === Now use scrollable_frame instead of enquiry_window for content ===

        # Number of Products Frame
        num_products_frame = ttk.Frame(scrollable_frame)
        num_products_frame.grid(row=0, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        ttk.Label(num_products_frame, text="Number of Products:").grid(row=0, column=0, padx=5, pady=5)
        num_products_entry = ttk.Entry(num_products_frame, width=10)
        num_products_entry.grid(row=0, column=1, padx=5, pady=5)
        num_products_entry.insert(0, "5")  # Default value

        # Products Frame
        products_frame = ttk.LabelFrame(scrollable_frame, text="Products")
        products_frame.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        def create_product_entries():
            # Clear existing entries
            for widget in products_frame.winfo_children():
                widget.destroy()

            try:
                num_products = int(num_products_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number of products")
                return

            # Labels
            ttk.Label(products_frame, text="Description").grid(row=0, column=0)
            ttk.Label(products_frame, text="Quantity").grid(row=0, column=1)
            ttk.Label(products_frame, text="Unit").grid(row=0, column=2)
            ttk.Label(products_frame, text="Last Price").grid(row=0, column=3)
            ttk.Label(products_frame, text="Price Unit").grid(row=0, column=4)

            # Entries
            self.product_entries = []
            for i in range(num_products):
                desc_entry = ttk.Entry(products_frame, width=30)
                desc_entry.grid(row=i+1, column=0, padx=5, pady=5)

                qty_entry = ttk.Entry(products_frame, width=10)
                qty_entry.grid(row=i+1, column=1, padx=5, pady=5)

                unit_entry = ttk.Entry(products_frame, width=10)
                unit_entry.grid(row=i+1, column=2, padx=5, pady=5)

                price_entry = ttk.Entry(products_frame, width=10)
                price_entry.grid(row=i+1, column=3, padx=5, pady=5)

                price_unit_entry = ttk.Entry(products_frame, width=10)
                price_unit_entry.grid(row=i+1, column=4, padx=5, pady=5)

                self.product_entries.append({
                    'description': desc_entry,
                    'quantity': qty_entry,
                    'unit': unit_entry,
                    'last_price': price_entry,
                    'price_unit': price_unit_entry
                })

        # Create Initial Entries
        create_product_entries()

        # Update Products Button
        ttk.Button(num_products_frame, text="Update Products",
                command=create_product_entries).grid(row=0, column=2, padx=5, pady=5)

        # Suppliers Selection Frame
        suppliers_frame = ttk.LabelFrame(scrollable_frame, text="Select Suppliers")
        suppliers_frame.grid(row=2, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        self.supplier_vars = []
        row, col = 0, 0
        for supplier, contacts in self.suppliers.items():
            supplier_label = ttk.Label(suppliers_frame, text=supplier, font=('Helvetica', 10, 'bold'))
            supplier_label.grid(row=row, column=col, sticky=tk.W, padx=5, pady=5)
            row += 1

            for contact in contacts:
                var = tk.BooleanVar()
                cb = ttk.Checkbutton(suppliers_frame,
                                    text=f"{contact[0]} ({contact[1]})",
                                    variable=var)
                cb.grid(row=row, column=col, sticky=tk.W, padx=15, pady=5)
                self.supplier_vars.append((var, supplier, contact))
                row += 1

            if row > 15:
                row = 0
                col += 1

        # Generate Emails Button
        ttk.Button(scrollable_frame, text="Generate Emails",
                command=self.generate_emails).grid(row=3, column=0, pady=10)
  
    def generate_emails(self):
        """Generate emails for selected suppliers"""
        # Collect products
        products = []
        for entry in self.product_entries:
            desc = entry['description'].get().strip()
            qty = entry['quantity'].get().strip()
            unit = entry['unit'].get().strip()
            last_price = entry['last_price'].get().strip()
            price_unit = entry['price_unit'].get().strip()
            
            if desc:  # Only add if description is not empty
                product_str = f"{desc} - {qty} {unit}"
                if last_price:
                    product_str += f" - last price paid: £{last_price}/{price_unit}"
                products.append(product_str)

        # Read email template
        with open(self.email_template_file, 'r') as f:
            email_template = f.read()

        # Generate emails for selected suppliers
        emails_window = self.create_toplevel("Generated Emails", "800x600")

        # Text widget to display emails
        email_text = tk.Text(emails_window, wrap=tk.WORD)
        email_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # List to store emails used
        used_emails = []

        # Generate emails for selected suppliers
        for var, supplier_name, contact_details in self.supplier_vars:
            if var.get():
                # Replace tokens in template
                email_content = email_template.replace('[contact_name]', contact_details[0])
                email_content = email_content.replace('[product_list]', '\n'.join(products))
                
                # Add to text widget
                email_text.insert(tk.END, f"--- Email to {contact_details[0]} from {supplier_name} ---\n")
                email_text.insert(tk.END, email_content + "\n\n")
                
                # Store contact details
                used_emails.append(f"{contact_details[0]} ({contact_details[1]}) - {supplier_name}")

        # Add list of used emails at the end
        email_text.insert(tk.END, "\n--- Contacts Included ---\n")
        for email in used_emails:
            email_text.insert(tk.END, f"{email}\n")

        # Add Copy All button
        def copy_all():
            pyperclip.copy(email_text.get("1.0", tk.END))
            messagebox.showinfo("Copied", "All emails copied to clipboard!")

        ttk.Button(emails_window, text="Copy All Emails", command=copy_all).pack(pady=10)

def main():
    root = tk.Tk()
    app = SteelEnquiryApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()