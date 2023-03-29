import openpyxl
import aiohttp
import asyncio
import os
import subprocess
import re

def replace_urls(filename, sheetname, url_pattern):
    wb = openpyxl.load_workbook(filename)
    sheet = wb[sheetname]
    
    url_counter = 0
    new_urls = []
    for row in sheet.iter_rows(min_row=2):
        cell = row[3]
        if cell.value:
            new_url = url_pattern.format(str(url_counter))
            new_urls.append(new_url)
            cell.value = new_url
            url_counter += 1
    
    wb.save(filename)

    return new_urls


async def download_image(session, url):
    async with session.get(url, allow_redirects=True) as response:
        image_index = re.findall(r'/id/(\d+)/', url)[0]
        image_filename = f"{image_index}.jpg"
        image_path = os.path.join('images', image_filename)
        with open(image_path, 'wb') as f:
            while True:
                chunk = await response.content.read(1024)
                if not chunk:
                    break
                f.write(chunk)
        return image_path


async def download_images(image_urls):
    async with aiohttp.ClientSession() as session:
  
        tasks = []
        for index, image_url in enumerate(image_urls):
            tasks.append(download_image(session, image_url))
        
  
        image_filenames = await asyncio.gather(*tasks)
    
    return image_filenames

def create_video(image_filenames):
    image_filenames = sorted(image_filenames, key=lambda x: int(os.path.splitext(os.path.basename(x))[0]))
    
    os.chdir('images/')
    
    subprocess.call(['ffmpeg', '-y', '-i', '%d.jpg', '-pix_fmt', 'yuv420p', '-r', '30', 'output.mp4'])
    
    for image_filename in image_filenames:
        os.remove(image_filename)

async def main():
    filename = 'xlsx/video.xlsx'
    sheetname = 'Respostas ao formulário 1'

    image_urls = replace_urls(filename, sheetname, 'https://picsum.photos/id/{}/1000/1000')
    
    image_filenames = await download_images(image_urls)
    create_video(image_filenames)

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())
