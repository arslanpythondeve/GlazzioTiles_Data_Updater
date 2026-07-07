import os
import requests


class GlazzioTilesImagesPipeline:
    def process_item(self, item, spider):
        folder = "output/images"
        os.makedirs(folder, exist_ok=True)

        image_names = []

        for i, url in enumerate(item.get("image_urls", []), start=1):
            filename = f"{item['Sku']}_{i}.jpg"
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