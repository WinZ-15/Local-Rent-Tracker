🏠 Local Rent Tracker
📌 About
A local desktop-based Rent Management System built using Python and Tkinter.
Designed to efficiently manage 1000+ tenants, it allows tracking of tenant details, monthly billing, utilities, and payment history.
The application focuses on speed, simplicity, and real-world usability for medium-scale rental management.
<br>
⚙️ Core Features
🧾 Tenant Management

Store tenant details:

Name
Block
Room


Organize records by Nepali Year and Month


💡 Utility Billing
Supports automatic charge calculation based on:

Water → Units × Rate
Electricity → Units × Rate
Internet → Units × Rate
Miscellaneous charges (manual input)


🧮 Auto Calculations
The system automatically calculates:

✅ Total Bill
✅ Paid Amount
✅ Outstanding Balance
✅ Advance Balance


🔄 Smart Balance System

Overpayment → Advance

Automatically carried forward and deducted next month


Underpayment → Outstanding

Added to next month's total




🖥️ UI & Design
🎨 Dashboard Style

Modern card-based layout
Displays 4 tenants per row
Fully scrollable interface
Color-coded cards:

🔴 Outstanding (due)
🟢 Advance (credit)
⚫ Neutral




🧩 Functionality
➕ Add Tenant

Form-based input system
Automatically:

Selects current month
Defaults to Baisakh if none selected




✏️ Edit Tenant

Opens detailed edit dialog
Supports:

Full field editing
Notes/comments
Credit number entry




⚡ Quick Inline Editing

Excel-like editing inside dashboard
Double-click cells to edit:

Units
Rates
Charges
Paid amount


Automatically recalculates totals


🔁 Carry Over System
Single Carry

Moves tenant to next month
Keeps:

Name, Block, Room
Rates


Resets usage
Automatically calculates:

Balance OR Advance



Batch Carry

Carry all tenants of a month to the next month
One-click operation


❌ Delete Tenant

Removes tenant record
Instantly updates UI and sidebar


🧭 Navigation & Search
📂 Sidebar Navigation

Organized by:

Year → Month


Click to instantly load tenants


🔍 Search System

Search tenants by name
Automatically:

Setting button
Change Global rent 
new Records will use the new rent 
old record remain the same
Highlights match
Scrolls to tenant
