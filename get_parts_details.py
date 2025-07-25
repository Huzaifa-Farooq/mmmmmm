from pywinauto import Application, mouse, Desktop
import pywinauto
import pandas as pd
import time
import os
import re
import mss
from PIL import Image
from concurrent.futures import ThreadPoolExecutor
from loguru import logger
import traceback
from functools import wraps
import math


def timeit(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        logger.info(f"▶️ Starting: {func.__name__}")
        result = func(*args, **kwargs)
        end_time = time.perf_counter()
        duration = end_time - start_time
        logger.info(f"✅ Finished: {func.__name__} in {duration:.4f} seconds")
        return result
    return wrapper



DONE_BID_FILE = "done_bids.txt"
REF_DATA = {}
IMAGE_CORDINATES = {}
clean_name = re.compile(r'\/|\#|\*|\?|\:|\"|\<|\>|\||\+|\=|\;|\@|\$|\,')


def mark_bid_as_done(bid):
    with open(DONE_BID_FILE, "a") as f:
        f.write(f"{bid}\n")


def load_crawled_section_pairs(filepath):
    if not os.path.exists(filepath):
        return []
    
    df = pd.read_csv(filepath)
    section_pairs = df[['U_Section', 'U_SubSection']].values.tolist()
    section_pairs = [tuple(pair) for pair in section_pairs]
    return section_pairs


def load_done_bids():
    try:
        with open(DONE_BID_FILE, "r") as f:
            return set(line.strip() for line in f)
    except FileNotFoundError:
        return set()
    

def load_bids():
    df = pd.read_excel("parts_data.xlsx")
    global REF_DATA
    REF_DATA = {
        i[0]: i[1:]
        for i in df[['U_ModelCode', 'U_Machine Type', 'U_Model', 'U_SGLUniqueModelCode']].values.tolist()
    }

    bids = df['U_ModelCode'].unique().tolist()
    return bids


def save_image(img_path, main_win):
    global IMAGE_CORDINATES

    if not os.path.exists(img_path):
        if not IMAGE_CORDINATES:
            img_elem = main_win.child_window(auto_id="frmViewerPicture").children(control_type="Pane")[0].children(control_type="Pane")[0]
            rect = img_elem.rectangle()
            IMAGE_CORDINATES = {
                "top": rect.top,
                "left": rect.left,
                "width": rect.width(),
                "height": rect.height()
            }
            bmp = img_elem.capture_as_image()
            bmp.save(img_path, "JPEG")  # Save the image
        else:
            with mss.mss() as sct:
                # Capture the image using the stored coordinates
                img = sct.grab(IMAGE_CORDINATES)
                bmp = Image.frombytes("RGB", img.size, img.rgb)
                bmp.save(img_path, "JPEG")
    # else:
    #     exit(f"Image already exists: {img_path}")


def get_clean_filename(section_text, sub_section_text, sgl_model_code, img_index=0):
    def clean(sec_val):
        sec_val = sec_val.strip()
        if ' / ' in sec_val:
            part_number = sec_val.split(' ')[0]
            sec_val = sec_val.split(' / ')[1].split('#')[0].strip()
            sec_val = '{} {}'.format(str(part_number), sec_val)
        sec_val = clean_name.sub('', sec_val)
        sec_val = re.sub(r'\s+', ' ', sec_val).strip()
        return sec_val[:65].replace(' ', '-').title()
    
    section_cleaned = clean(section_text)
    sub_section_cleaned = clean(sub_section_text)

    if img_index:
        return '{}-{}-{}-{}.jpg'.format(
            sgl_model_code,
            section_cleaned,
            sub_section_cleaned,
            str(img_index)
        )
    else:
        return '{}-{}-{}.jpg'.format(
            sgl_model_code,
            section_cleaned,
            sub_section_cleaned
        )
    
@timeit
def extract_table_data(main_win):
    logger.info("Extracting table data...")
    parts_list_table_e = main_win.child_window(
        control_type='Table', 
        class_name="WindowsForms10.Window.8.app.0.33c0d9d", 
        auto_id='dgPartslist',
        title="DataGridView"
        ).wrapper_object()
    
    table_childs = parts_list_table_e.children()
    if not table_childs:
        time.sleep(1)
        table_childs = parts_list_table_e.children()

    # TODO - REMOVE
    if not table_childs:
        exit("No table data found, exiting...")

    header_row = next(iter(i for i in table_childs if "Top Row" in i.window_text()), None)
    headers = [h.window_text() for h in header_row.children() if h.element_info.control_type == "Header"][1:5]

    parts_data = []
    for row in [i for i in table_childs if i.element_info.control_type not in ["ScrollBar"] and "Top Row" not in i.window_text()]:
        row_data = [cell.iface_value.CurrentValue for cell in row.children()[1:5]]
        parts_data.append({i[0]: i[1] for i in zip(headers, row_data)})

    # TODO - REMOVE
    if not parts_data:
        exit("No parts data found, exiting...")

    logger.info(f"Extracted {len(parts_data)} parts from the table.")
    return parts_data

@timeit
def get_next_list_button(main_win):
    logger.info("Getting next list button...")
    next_button = main_win.child_window(title="Next List", control_type="Button").wrapper_object()
    
    is_enabled = next_button.is_enabled()
    logger.info(f"Got next list btn")
    return next_button, is_enabled

@timeit
def get_next_picture_button(main_win):
    logger.info("Getting next picture button...")
    next_picture_button = main_win.child_window(auto_id="frmViewerPicture").child_window(
        title="Next Picture", control_type="Button").wrapper_object()
    
    is_enabled = next_picture_button.is_enabled()

    logger.info(f"Got next picture button")
    return next_picture_button, is_enabled


def get_parts_details(bid):
    global REF_DATA
    machine_type, model, sgl_model_code = REF_DATA[bid]
    csv_filepath = f"parts_database/{bid}.csv"
    
    crawled_section_pairs = load_crawled_section_pairs(csv_filepath)

    mouse.move(coords=(1, 1))
    # Connect to the app
    try:
        app = Application(backend="uia").connect(title_re=".*KUBOTA.*BKID.*")
    except Exception as e:
        time.sleep(2)
        app = Application(backend="uia").connect(title_re=".*KUBOTA.*BKID.*")

    # Access the main window by automation ID
    main_win = app.window(auto_id="frmViewer")

    # sometimes we have structure like this Top Section -> Section -> Sub Section
    # In other cases it is Section -> Sub Section
    main_sec_e = main_win.child_window(auto_id="tvBook")
    main_sec_child_trees = main_sec_e.children(control_type="TreeItem")
    main_tree = main_sec_child_trees[0]
    sections = main_tree.children(control_type="TreeItem")
    if not sections[0].children() and not sections[0].children():
        with open('two_level.txt', 'a') as f:
            f.write(str(bid) + '\n')

        sections = main_sec_child_trees

    for section in sections:
        section_text = section.window_text()
        logger.info(f"Section: {section_text}")
        section.expand()

        # Access the Sub-Section tree items
        sub_sections = section.children(control_type="TreeItem")
        for i in range(5):
            if not sub_sections:
                sub_sections = section.children(control_type="TreeItem")
                time.sleep(0.2)

        time.sleep(0.1)
        
        for sub_section in sub_sections:
            sub_section_text = sub_section.window_text()
            logger.info(f"  Sub-Section: {sub_section_text}")
            if (section_text, sub_section_text) in crawled_section_pairs:
                logger.info(f"Skipping already crawled section: {section_text} - {sub_section_text}")
                continue

            try:
                sub_section.expand()
            except:
                pass

            dataset = []
            parts_data = []
            img_index = 0
            # keep going to next images and getting parts until
            while True:
                next_image_button, is_image_btn_enabled = None, None
                # keep getting parts table data until the "Next List" button is enabled
                while True:
                    try:
                        table_parts_data = extract_table_data(main_win)
                    except pywinauto.findwindows.ElementNotFoundError as ex:
                        if 'frequently used items' in section_text.lower():
                            table_parts_data = []
                        else:
                            raise ex

                    if not table_parts_data and 'frequently used items' in section_text.lower():
                        next_button, is_btn_enabled = None, None
                    else:
                        next_button, is_btn_enabled = get_next_list_button(main_win)
                        
                    # with ThreadPoolExecutor(max_workers=3) as executor:
                    #     future_table = executor.submit(extract_table_data, main_win)
                    #     # Check if the "Next List" button is enabled
                    #     future_button = executor.submit(get_next_list_button, main_win)
                    #     # check if the "Next Picture" button is enabled only if it's None
                    #     if next_image_button is None and  is_image_btn_enabled is None:
                    #         future_image_button = executor.submit(get_next_picture_button, main_win)
                    #     else:
                    #         future_image_button = None

                    #     # Wait all to finish
                    #     table_parts_data = future_table.result()
                    #     next_button, is_btn_enabled = future_button.result()
                    #     if future_image_button:
                    #         next_image_button, is_image_btn_enabled = future_image_button.result()

                    parts_data.extend(table_parts_data)

                    # TODO - REMOVE
                    if is_btn_enabled:
                        exit("Next button is enabled")
                        break
                        next_button.click()
                        time.sleep(1)
                    
                    if not is_btn_enabled:
                        break

                img_filename = get_clean_filename(section_text, sub_section_text, sgl_model_code, img_index)
                img_path = os.path.join("images", img_filename)
                save_image(img_path, main_win)
                img_index += 1

                logger.info(f"{len(parts_data)} parts found in this section")
                data = [
                    {
                        'Machine Type': machine_type,
                        'Model': model, 
                        'BKID': bid,
                        'U_SGLUniqueModelCode': sgl_model_code,
                        'U_Section': section_text,
                        'U_SubSection': sub_section_text,
                        'U_SectionDiagram': os.path.basename(img_path),
                        **part_data
                    }
                    for part_data in parts_data
                ]
                dataset.extend(data)

                next_image_button, is_image_btn_enabled = get_next_picture_button(main_win)

                # Check if there is a "Next Picture" button to go to the next image
                if is_image_btn_enabled:
                    logger.info("Moving to next image...")
                    next_image_button.click()
                    time.sleep(1)
                else:
                    logger.info("No more images to process in this section.")
                    break

            df = pd.DataFrame(dataset)
            df.to_csv(csv_filepath, index=False, mode='a',
                      header=not os.path.exists(csv_filepath))
    
    app.kill()


def navigate_to_bid(bid, instance_index):
    app_path = r"C:/Program Files (x86)/KLTD/GSPcLocal/GSPcLocalViewer.exe"
    app = Application(backend="uia").start(cmd_line=app_path)
    time.sleep(5)
    # Access the main window by automation ID
    main_win = app.window(auto_id="frmViewer")
    section = main_win.child_window(auto_id="tvBrowse")
    
    main_tree = section.child_window(title="KUBOTA-PAD KDG", control_type="TreeItem").child_window(title="KUBOTA_PAD", control_type="TreeItem")
    main_tree.expand()

    try:
        main_tree = section.child_window(title="KUBOTA_PAD", control_type="TreeItem")
    except Exception as e:
        time.sleep(5)
        main_tree = section.child_window(title="KUBOTA_PAD", control_type="TreeItem")

    categories = main_tree.children(control_type="TreeItem")
    if int(instance_index) > 2:
        categories = list(reversed(categories))

    for category in categories:
        category_text = category.window_text()
        logger.info(f"Category: {category_text}")
        category.expand()
        time.sleep(0.1)

        # Access the Sub-Category tree items
        sub_categories = category.children(control_type="TreeItem")
        for sub_category in sub_categories:
            sub_category_text = sub_category.window_text()
            logger.info(f"  Sub-Category: {sub_category_text}")
            sub_category.expand()
            time.sleep(0.1)
            
            # Access and logger.info all items under the Sub-Category
            items = sub_category.children(control_type="TreeItem")
            for item in items:
                item_text = item.window_text()
                if item_text == bid:
                    try:
                        item.expand()
                    except:
                        pass
                    # double click item
                    item.double_click_input(button='left')
                    return True


def kill_app():
    try:
        app = Application(backend="uia").connect(title_re=".*KUBOTA.*")
        app.kill()
    except:
        pass

    try:
        app = Application(backend="uia").connect(title_re=".*Viewer.*")
        app.kill()
    except:
        pass


INSTANCE_INDEX = os.getenv("INSTANCE_INDEX")

kill_app()
os.makedirs("images", exist_ok=True)
os.makedirs("parts_database", exist_ok=True)

all_bids = load_bids()
# divide bids to 5 batches and operate on INSTANCE_INDEX
chunk_size = math.ceil(len(all_bids) / 5)
bids_patches = [all_bids[i:i + chunk_size] for i in range(0, len(all_bids), chunk_size)]
bids = bids_patches[int(INSTANCE_INDEX)]

completed_bids = load_done_bids()
bids = [bid for bid in bids if bid not in completed_bids]
if int(INSTANCE_INDEX) > 2:
    bids = list(reversed(bids))

logger.info(f"{len(bids)} BookIDs remaining...")

for bid in bids:
    logger.info(f"At bid: {bid}")
    try:
        for i in range(5):
            try:
                r = navigate_to_bid(bid, INSTANCE_INDEX)
                break
            except Exception as e:
                logger.error(f"Error while navigating to bid. {str(e)}")
                kill_app()
                continue
        if r:
            get_parts_details(bid)
            mark_bid_as_done(bid)
        else:
            logger.info(f"Failed to navigate to {bid}")
            kill_app()
    except Exception as e:
        err = traceback.format_exc()
        logger.info(f"Error: {err}")
        exit()
        # kill_app()
