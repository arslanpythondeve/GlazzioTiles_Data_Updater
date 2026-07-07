import os, json
import requests


class GlazzioTilesImagesPipeline:
    def __init__(self):
        self.config = self.read_json_file('input/config.json')
        self.input_sku_column_name = self.config.get('excel_sku_column_name', '')

    def process_item(self, item, spider):
        sku = '_' + self.input_sku_column_name
        folder = "output/images"
        os.makedirs(folder, exist_ok=True)

        image_names = []

        for i, url in enumerate(item.get("image_urls", []), start=1):
            filename = f"{item.get(sku, 'skuNoFound')}_{i}.jpg"
            filepath = os.path.join(folder, filename)

            try:
                r = requests.get(url, timeout=30)
                if r.status_code == 200:
                    with open(filepath, "wb") as f:
                        f.write(r.content)

                    image_names.append(filename)

            except Exception as e:
                spider.logger.error(f"Image download failed: {url} - {e}")

        item["images_name"] = ", ".join(image_names)

        return item

    def read_json_file(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)

        except Exception as e:
            print(f"Error reading JSON file: {e}")

        return {}
