from selenium.webdriver.common.by import By


class Selectors:
    LnbFreq = (By.ID, "LnbFreq")
    SateFreq = (By.ID, "SateFreq")
    SateSr = (By.ID, "SateSr")
    
    Encoder1_btn = (By.ID, "Encoder1_output")
    Encoder2_btn = (By.ID, "Encoder2_output")
    Encoder3_btn = (By.ID, "Encoder3_output")
    Encoder4_btn = (By.ID, "Encoder4_output")

    Encoder1 = (By.CSS_SELECTOR, "[id^='Encoder1_']")
    Encoder2 = (By.CSS_SELECTOR, "[id^='Encoder2_']")
    Encoder3 = (By.CSS_SELECTOR, "[id^='Encoder3_']")
    Encoder4 = (By.CSS_SELECTOR, "[id^='Encoder4_']")
