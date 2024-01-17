# Noslēguma projekta darbs priekšmetā "Lietojumprogrammatūras automatizēšanas rīki"
## Projekta uzdevums
Izveidot Python programmu, kuru izpildot iegūst jaunākos dzīvokļu sludinājumus ss.com mājaslapā pēc iepriekš norādītiem kritērijiem, kas atrodas Excel failā, kur arī beigās sludinājumi tiek saglabāti. Šī programma automatizē, paātrina un vienkāršo veidu, kā tās lietotājs var apskatīt jaunākos dzīvokļu sludinājumus pēc iepriekš izvēlētiem kritērijiem, jo lietotājam nav katru reizi manuāli jāapmeklē sludinājumu portāls un jāatliek nepieciešami filtri.
## Projektā izmantotās Python bibliotēkas
### openpyxl:
Šī bibliotēka ļauj lasīt un rakstīt Excel failus. Šajā projektā tā tiek izmantota, lai manipulētu ar Excel datiem, proti, lasīt sludinājumu filtru vērtības un ierakstīt iegūtos sludinājumus atpakaļ Excel failā. Bibliotēka tika izvēlēta, jo tā nodrošina vienkāršu un efektīvu veidu, kā strādāt ar Excel datiem Python vidē.
### selenium:
Tīmekļa automatizācijas rīks, kas ļauj veikt tīmekļa pārlūkprogrammu darbības. Projektā tas tiek izmantots, lai atvērtu mājas lapu, ievadītu datus meklēšanas formās un iegūtu nepieciešamo informāciju no sludinājumu portāla. Šis rīks tika izmantots, jo tas ļauj pilnībā automatizēt tīmekļa pārlūkošanas un datu ieguves procesu, kā arī piedāvā plašas iespējas mijiedarboties ar tīmekļa lapas elementiem.
### time:
Šī ir Python standarta bibliotēka, kas tiek izmantota laika manipulācijām. Projektā tā tiek izmantota, lai ieviestu laika aizkaves, kas ir nepieciešamas, lai nodrošinātu, ka tīmekļa lapas ir pilnībā ielādētas pirms datu ieguves.
## Programmatūras apraksts
Excel faila ielāde, darba lapu izvēle un WebDriver konfigurācija:
```
wb = load_workbook('dzivokli.xlsx')
ws = wb['Filtrs']
ws2 = wb['Sludinajumi']
max_row = ws.max_row
service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)
```
Šī programmas daļa izmanto Python ciklu, lai iegūtu datus no Excel faila darba lapas 'Filtrs'. Šeit tiek izgūtas vērtības no konkrētām šūnām katrā rindā un piešķirtas mainīgajiem:
```
for row in range(2, max_row+1):
    lokacija = (ws['a' + str(row)].value)
    max_cena = (ws['b' + str(row)].value)
    dar_veids = (ws['c' + str(row)].value)
```
Tīmekļa lapas manipulācija, izmantojot Selenium WebDriver. Šī programmas daļa nodrošina mājas lapas atvēršanu, datu ievadīšanu meklēšanas formā un meklēšanas darbības uzsākšanu:
```
url = "https://www.ss.com/lv/real-estate/flats/" + lokacija
driver.get(url)

# Izvēlas darījuma veidu
select_element = driver.find_element(By.NAME, "sid")
select = Select(select_element)
select.select_by_visible_text(dar_veids)

# Uzliek sludinājumu maksimālo cenu
element = driver.find_element(By.ID, "f_o_8_max")
element.send_keys(max_cena)

# Nospiež pogu meklēt
meklet = driver.find_element(
    By.CSS_SELECTOR, "input[type='submit'][value='Meklēt']")
meklet.click()

time.sleep(5)
```
Šī programmas koda daļa ir veltīta sludinājumu datu iegūšanai no tīmekļa lapas, izmantojot Selenium WebDriver. Tiek izmantots XPath, lai atrastu visas tabulas rindas (<tr> elementus) tīmekļa lapā, kuru ID satur 'tr_'. Pēc tam tiek iegūta konkrēta informāciju no katra objekta (dzīvokļa atrašānās iela, dzīvokļa platība, cena, links uz sludinājumu):
```
rows = driver.find_elements(
    By.XPATH, "//tr[contains(@id, 'tr_')]")[:15]
sludinajumi = []
for row in rows:
    iela = row.find_element(
        By.XPATH, ".//td[contains(@class, 'msga2-o')][1]/a").text

    m2 = row.find_element(
        By.XPATH, ".//td[contains(@class, 'msga2-o')][3]/a").text
    cena = row.find_element(
        By.XPATH, ".//td[contains(@class, 'msga2-o')][last()]/a").text

    link = row.find_element(
        By.XPATH, ".//td[contains(@class, 'msg2')]/div/a").get_attribute('href')

    sludinajumi.append(
        {"iela": iela, "m2": m2, "cena": cena, "link": link})
```
Sludinājumu datu ierakstīšana Excel failā, tā saglabāšana un beigās tiek nodrošināta tīmekļa pārlūkprogrammas instances aizvēršana:
```
row = 2

# Atlasītos sludinājumus saglabā Excel failā
for s in sludinajumi:
    ws2.cell(row=row, column=1, value=s['iela'])
    ws2.cell(row=row, column=2, value=int(s['m2']))
    ws2.cell(row=row, column=3, value=s['cena'])
    ws2.cell(row=row, column=4, value=s['link'])
    row += 1

wb.save('dzivokli.xlsx')
driver.close()
```
