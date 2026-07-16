import os
import requests
import json

PEXELS_API_KEY = "XOjgxKTCF7CSTaFQJinfTfoG7Zro51rTdFiYTuJ3BeCUK1Tl6tBegfzf"

SCENE_QUERIES = {
    "Video_1": {
        "1.1": "university library study room",
        "2.1": "indian red spice market",
        "2.2": "farmers market fresh produce",
        "2.3": "restaurant dinner table",
        "3.1": "handloom weaving loom",
        "3.2": "green organic farming agriculture",
        "4.1": "business boardroom meeting",
        "5.1": "technology office workspace",
        "6.1": "supermarket retail store shelves",
        "7.1": "steel factory industrial plant",
        "8.1": "modern university library building"
    },
    "Video_2": {
        "1.1": "academic library bookshelves",
        "2.1": "passenger boarding car using smart phone app",
        "3.1": "modern hospital lobby reception desk",
        "3.2": "spicy rice biryani plate",
        "4.1": "classroom student training coaching",
        "4.2": "hotel front desk concierge check in",
        "5.1": "professor classroom teaching students"
    },
    "Video_3": {
        "1.1": "world map globe background",
        "2.1": "cargo ship shipping container harbor",
        "3.1": "pizza delivery driver motorcycle rider",
        "4.1": "fast food burger menu board restaurant",
        "4.2": "multicultural office business meeting",
        "5.1": "university graduation stage graduation cap"
    }
}

def get_pexels_image(query, api_key):
    url = "https://api.pexels.com/v1/search"
    headers = {"Authorization": api_key}
    params = {
        "query": query,
        "per_page": 1,
        "orientation": "landscape"
    }
    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        if response.status_code == 200:
            data = response.json()
            photos = data.get("photos", [])
            if photos:
                # Return the landscape medium or large size URL
                return photos[0]["src"]["medium"]
        else:
            print(f"Pexels API Error for query '{query}': Status {response.status_code}")
    except Exception as e:
        print(f"Network error querying Pexels: {e}")
    return None

def download_image(url, save_path):
    try:
        response = requests.get(url, stream=True, timeout=15)
        if response.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            print(f"Downloaded: {save_path}")
            return True
    except Exception as e:
        print(f"Failed to download image {url}: {e}")
    return False

def generate_pexels_storyboard():
    master_dir = "/Users/rkanduk/Documents/GitHub/OpenMontage/professor_storyboard"
    os.makedirs(master_dir, exist_ok=True)
    
    # Mapping for folder names
    folder_mapping = {
        "Video_1": "Video_1_Foundations",
        "Video_2": "Video_2_Services",
        "Video_3": "Video_3_Global"
    }
    
    for video_key, scenes in SCENE_QUERIES.items():
        subfolder = folder_mapping[video_key]
        video_dir = os.path.join(master_dir, subfolder)
        os.makedirs(video_dir, exist_ok=True)
        
        for scene_num, query in scenes.items():
            filename = f"{video_key}_Scene_{scene_num}_Slide.png"
            dest_path = os.path.join(video_dir, filename)
            
            print(f"Searching Pexels for: '{query}'...")
            img_url = get_pexels_image(query, PEXELS_API_KEY)
            
            if img_url:
                download_image(img_url, dest_path)
            else:
                print(f"Could not find Pexels image for query: '{query}'")

if __name__ == "__main__":
    generate_pexels_storyboard()
