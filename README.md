# InvMan - Inventory Manager 📦

<div id="header" align="left">
 <img src="headerImg.png">
</div>

InvMan is a comprehensive inventory management application built with PyQt5. It helps businesses manage their inventory, track performance, and process orders and refunds efficiently.

## Features ✨

- Add, edit, and remove inventory items
- Track stock levels by item and model
- Process orders and refunds
- Apply discounts to orders
- Track performance over a specified date range
- Filter and search inventory by name, phone model, and category
- User-friendly interface with image support

## Installation ⚙️

### From Executable 💻

To run InvMan from the executable:

1. Download the latest version of `InvMan.rar` from the [releases page](https://github.com/dizzydroid/InvMan/releases).
2. Extract and run the executable `main.exe` located in the `dist` directory.
3. The `dist` folder will hold all the generated files in .xlsx format.

### From Source 🛠️

To install and run InvMan from source, follow these steps:

1. Clone the repository:
    ```bash
    git clone https://github.com/dizzydroid/InvMan.git
    cd InvMan
    ```

2. Create and activate a virtual environment (optional but recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

4. Run the application:
    ```bash
    python main.py
    ```

## Dependencies 📚

If you are running from source, the following dependencies are required:

- Python
- PyQt5
- pandas
- openpyxl

## Usage 📖

Upon launching the app, you will be presented with a search bar and filter options to easily manage your inventory. You can add new items, edit existing ones, process orders and refunds, and track performance over time.

### Adding a New Item ➕

1. Click the "Add Item" button.
2. Enter the item details, including name, category, image, and models.
3. Click "Add Item" to save.

### Processing an Order 🛒

1. Click on an item to open its options.
2. Click "Order" and fill in the order details.
3. Click "Order" to generate a receipt and update the inventory.

### Tracking Performance 📈

1. Click the "Track Performance" button.
2. Select the start and end dates.
3. Click "Track Performance" to view the net profit for the selected period.

## Found Bugs 🐞

If you encounter any bugs, please report them by creating an issue on the [GitHub Issues](https://github.com/dizzydroid/InvMan/issues) page.

## Contributing 🤝

Contributions are welcome! Please fork the repository and create a pull request with your changes.

## License 📄

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements 🙏

Thanks to the open-source community for providing the tools and libraries that made this project possible.
