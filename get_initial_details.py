from pywinauto import Application
import pandas as pd
import time


# Connect to the app
app = Application(backend="uia").connect(title_re=".*KUBOTA.*|.*GSPcLocal Viewer.*")  # Or use process=5032

# Access the main window by automation ID
main_win = app.window(auto_id="frmViewer")
section = main_win.child_window(auto_id="tvBrowse")

main_tree = section.child_window(title="KUBOTA-PAD KDG", control_type="TreeItem").child_window(title="KUBOTA_PAD", control_type="TreeItem")

# Now we will iterate over the tree childrens that are also of type tree.
# Then we click each child Sub-Category of the Category
# then we print text of all childrens of Sub-Category
dataset = []
for category in main_tree.children(control_type="TreeItem"):
    category_text = category.window_text()
    print(f"Category: {category_text}")
    category.expand()
    time.sleep(0.5)

    # Access the Sub-Category tree items
    sub_categories = category.children(control_type="TreeItem")
    for sub_category in sub_categories:
        sub_category_text = sub_category.window_text()
        print(f"  Sub-Category: {sub_category_text}")
        sub_category.expand()
        time.sleep(0.5)
        
        # Access and print all items under the Sub-Category
        items = sub_category.children(control_type="TreeItem")
        for item in items:
            item_text = item.window_text()
            data = {
                'U_Brand': 'Kubota',
                'U_Category': category_text,
                'U_Model': sub_category_text,
                'U_ModelCode': item_text,
            }
            dataset.append(data)
        
        sub_category.collapse()
    
    category.collapse()  # Collapse the category after processing

df = pd.DataFrame(dataset)
# Save the dataset to a JSON file
df.to_excel("kubota_data.xlsx", index=False)
