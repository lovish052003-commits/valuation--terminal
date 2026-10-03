import requests

r = requests.get('http://127.0.0.1:5000/')
assert 'sliderWacc" class="range-slider" min="5.0" max="25.0"' in r.text, 'WACC slider range check failed'
assert 'sliderGrowth" class="range-slider" min="0.0" max="50.0"' in r.text, 'Growth slider range check failed'
assert 'sliderTerminal" class="range-slider" min="1.0" max="10.0"' in r.text, 'Terminal slider range check failed'
assert 'sliderTax" class="range-slider" min="15.0" max="35.0" step="0.5" value="30.0"' in r.text, 'Tax slider range check failed'
assert 'badgeTax">30.0%' in r.text, 'Tax badge check failed'
assert 'headerCompanyLogo' not in r.text, 'headerCompanyLogo still in HTML'
assert 'dupontCompanyLogo' not in r.text, 'dupontCompanyLogo still in HTML'
assert 'altmanCompanyLogo' not in r.text, 'altmanCompanyLogo still in HTML'
print("ALL HTML AND SLIDER ASSERTIONS PASSED PERFECTLY!")
