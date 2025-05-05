from BudgetUtils import load_categories, save_categories

def edit_categories():
    """Prompt user to edit category names, allowing modifications, deletions, and skipping."""
    categories = load_categories()

    if not categories:
        print("\n\U0000274C categories.csv file not found or invalid.\n")
        return

    updated_categories = categories.copy()  # Start with a copy!
    names_to_remove = set()

    print("\n\U00002728 Edit Transaction Names and Categories\n")
    print("\U0001F4A1 Press Enter to keep, 'd' to delete, or 's' to skip to the end.")
    print("\U0001F4A1 Editing format: name,category (leaving blank before/after comma leaves the corresponding value unchanged)\n")

    for name, category in list(categories.items()):
        if name in names_to_remove: # Skip already marked for deletion
            continue

        user_input = input(f"Edit entry: [{name}, {category}] ").strip()

        if user_input.lower() == "s":
            print("\n\U000023E9 Skipping to deletion confirmation...")
            break  # Just break; updated_categories is already correct

        if user_input.lower() == "d":
            names_to_remove.add(name)
        elif user_input:
            try:
                parts = [x.strip() for x in user_input.split(",", 1)]
                new_name = (parts[0] if parts[0] else name).strip().upper()
                new_category = (parts[1] if len(parts) > 1 and parts[1] else category).strip().capitalize()

                if new_name in updated_categories and new_name != name and new_name not in names_to_remove: # Check against updated_categories
                    print(f"\U00002757 Duplicate name detected: {new_name}. Retaining original entry.")
                else:
                    if new_name != name:
                        del updated_categories[name] # Remove old entry from updated_categories
                    updated_categories[new_name] = new_category
            except ValueError:
                print("\U0000274C Invalid format! Please use 'Name,Category' format.")
        # No else needed; if empty, the original is kept in updated_categories

    # Remove deleted entries *after* the loop
    for name in names_to_remove:
        if name in updated_categories: #Check if still in updated_categories
            del updated_categories[name]

    if names_to_remove:
        print("\n\U0001F6A7 The following entries will be permanently removed from categories.csv:")
        for name in names_to_remove:
            print(f"- {name}, {categories[name]}")
        confirm = input("\nAre you sure you want to proceed? (y/n): ").strip().lower()
        if confirm != "y":
            print("\n\U0000274C No changes were made to the categories.")
            return

    save_categories(cats=updated_categories)
    print("\n\U00002705 Categories file has been cleaned and updated!\n")

if __name__ == "__main__":
    edit_categories()