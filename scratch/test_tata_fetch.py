import os, sys
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__) + '/..'))
import requests
from io import BytesIO
from PIL import Image
import company_logo_manager as clm

# Test fetching Tata Steel logo via DDG / Domain
logo = clm.fetch_company_logo("Tata Steel", "TATASTEEL", "http://www.tatasteel.com")
print("Fetched logo:", logo.size, logo.mode)
logo.save("scratch/tatasteel_fetched.png")
print("Saved scratch/tatasteel_fetched.png")
