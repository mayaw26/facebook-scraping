# this is the current version

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
import re
import pandas as pd

# Create a Facebook account to use for the scraping
fb_username = ""
fb_password = ""

def clean_text(text):
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_post_text(post_element):
    text_parts = []
    for element in post_element.find_elements(By.XPATH, ".//*"):
        if element.tag_name.lower() in ['span', 'div', 'p', 'li', 'a']:
            text = element.text.strip()
            if text:
                text_parts.append(text)
    full_text = ' '.join(text_parts)
    cleaned_text = clean_text(full_text)
    return cleaned_text

def is_post_valid(post_text):
    invalid_keywords = ["Page", "Video description", "Follow", "Share", "Comment"]
    return not any(keyword in post_text for keyword in invalid_keywords)

def scrape_posts(driver, profile_url, post_limit):
    driver.get(profile_url)
    time.sleep(5)

    # edit as HTML code on Facebook changes
    post_xpath = "//div[contains(@class, 'x11i5rnm') and contains(@class, 'xat24cr') and contains(@class, 'x1mh8g0r') and contains(@class, 'x1vvkbs') and contains(@class, 'x126k92a')]"
    scraped_posts = set()
    unique_posts = []
    while len(unique_posts) < post_limit:
        posts = driver.find_elements(By.XPATH, post_xpath)
        if not posts:
            print(f"No posts found or the page structure might have changed for {profile_url}.")
            break

        for post in posts:
            post_content = extract_post_text(post)
            if post_content and len(post_content) > 20:
                if is_post_valid(post_content):
                    if post_content not in scraped_posts:
                        scraped_posts.add(post_content)
                        unique_posts.append(post_content)
                        if len(unique_posts) >= post_limit:
                            break

        # Scroll down to load more posts
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)

    return unique_posts

def export_to_excel(profiles_posts):
    max_posts = 20
    main_data = []
    with pd.ExcelWriter('facebook_posts.xlsx', engine='xlsxwriter') as writer:
        # Create the main sheet
        for profile_name, (profile_url, posts) in profiles_posts.items():
            num_posts = len(posts)
            row = {'Facebook Account': profile_name, 'Profile URL': profile_url}
            row.update({f'Post {i+1}': posts[i] for i in range(num_posts)})
            for i in range(num_posts, max_posts):
                row[f'Post {i+1}'] = ''
            main_data.append(row)
        
        df_main = pd.DataFrame(main_data)
        df_main.to_excel(writer, sheet_name='Main', index=False)

        # Create individual sheets for each profile
        for profile_name, (profile_url, posts) in profiles_posts.items():
            df_posts = pd.DataFrame(posts, columns=['Post'])
            df_posts.to_excel(writer, sheet_name=profile_name, index=False)

    print(f"Data exported to facebook_posts.xlsx")

def main():
    try:
        chrome_options = Options()
        prefs = {"profile.default_content_setting_values.notifications": 2}
        chrome_options.add_experimental_option("prefs", prefs)

        service = Service('...//chromedriver.exe')
        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.get("http://www.facebook.com")
        username = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='email']")))
        password = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='pass']")))
        username.clear()
        username.send_keys(fb_username)
        password.clear()
        password.send_keys(fb_password)
        button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))).click()
        time.sleep(10)

        # Define the URLs to every profile that should be scraped from
        profiles = []

        post_limit = 20

        profiles_posts = {}
        for profile_name, profile_url in profiles:
            print(f"Scraping posts from {profile_url}...")
            posts = scrape_posts(driver, profile_url, post_limit)
            profiles_posts[profile_name] = (profile_url, posts)

        export_to_excel(profiles_posts)

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
