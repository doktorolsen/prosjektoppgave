#
# Prosjektoppgave, PY1010 USN
# Support dashboard
#

#%% del a - innlesing av xls til array

import pandas as pd
file = 'support_uke_24.xlsx' # oppretter variabel med filnavnet som skal leses inn

df = pd. read_excel ( file ) # leser fil til dataframe

# laster inn de ulike kolonnene i hver sin NumPy array
u_dag = df['Ukedag'].to_numpy()
kl_slett = df['Klokkeslett'].to_numpy()
varighet = df['Varighet'].to_numpy()
score = df['Tilfredshet'].to_numpy()


#%% del b - uthenting av antall henvendelser per ukedag

import matplotlib.pyplot as plt

# teller antall tilfeller av hver ukedag i dataframen (henvendelser per ukedag), tar for gitt at det er lov å jobbe mot df siden noe annet ikke er spesifisert
ukedag_antall = df.iloc[:, 0].value_counts()

# setter riktig rekkefølger på ukedagene
uke_rekkefølge = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag']
ukedag_antall= ukedag_antall.reindex(uke_rekkefølge)

# lager stolpediagram av tallene
plot = ukedag_antall.plot(kind='bar', color='black', figsize=(8, 6))
plot.grid(True, linestyle='--', alpha=0.7)
plt.title('Antall henvendelser for hver ukedag')
plt.xlabel('Ukedag')
plt.ylabel('Antall')
plt.xticks(rotation=45)
plt.show()

print(ukedag_antall)


#%% del c - uthenting av korteste og lengste samtaletid

korteste_samtaletid = min(varighet)
lengste_samtaletid = max(varighet)

print('Korteste og lengste samtaler for uke 24:\n')
print(f'Korteste samtaletid: {korteste_samtaletid}')
print(f'Lengste samtaletid: {lengste_samtaletid}')


#%% del d - beregning av gjennomsnittlig samtaletid for uken

# oppretter variabler som brukes i løkken
sum_sekunder = 0
teller = 0

# løkke som går gjennom hver samtale-varighet
for varighet_str in varighet:
    # splitter stringen til timer, minutter og sekunder ved hjelp av :-tegnet
    tid_deler = varighet_str.split(':')
    h = int(tid_deler[0])  # gjør timer om til en int
    m = int(tid_deler[1])  # gjør minutter om til en int
    s = int(tid_deler[2])  # gjør sekunder om til en int
    
    # gjør om til sekunder og summerer
    sum_sekunder += h * 3600 + m * 60 + s
    teller += 1  # inkrementerer telleren for å gå til neste oppføring

# finner gjennomsnittet i sekunder
snitt_sekunder = sum_sekunder / teller

# gjør om sekundene til det opprinnelige '00:00:00' formatet ved hjelp av heltallsdivisjon og rest
snitt_tid = f'{int(snitt_sekunder // 3600):02}:{int((snitt_sekunder % 3600) // 60):02}:{int(snitt_sekunder % 60):02}'
print(f'Gjennomsnittlig samtaletid for alle samtaler i uke 24: {snitt_tid}')


#%% del e - uthenting av antall henvendelser per 2-timers bolk

from datetime import datetime #importerer datetime-modulen

# funksjon som teller henvendelser basert på de definerte totimersbolkene
def ant_henv_per_bolk(henvser):
    
    bolker = { # oppretter en dictionary for bolkene, key: string/navn for bolken value: tuple med datetime-objekt for start og slutt
        "08:00-10:00": (datetime.strptime("08:00:00", "%H:%M:%S"), datetime.strptime("10:00:00", "%H:%M:%S")),
        "10:00-12:00": (datetime.strptime("10:00:00", "%H:%M:%S"), datetime.strptime("12:00:00", "%H:%M:%S")),
        "12:00-14:00": (datetime.strptime("12:00:00", "%H:%M:%S"), datetime.strptime("14:00:00", "%H:%M:%S")),
        "14:00-16:00": (datetime.strptime("14:00:00", "%H:%M:%S"), datetime.strptime("16:00:00", "%H:%M:%S")),
    }

    antall_dict = {bolk: 0 for bolk in bolker} # oppretter dictionary for å telle henvendelser per bolk, alle bolkene settes til å begynne på 0
    
    # løkke for å gå gjennom henvendelsene og telle dem i riktig bolk
    for henv in henvser:
        henv_tid = datetime.strptime(henv, "%H:%M:%S")
        for bolk, (start, end) in bolker.items():
            if start <= henv_tid < end:
                antall_dict[bolk] += 1
                break

    return antall_dict # returnerer dictionary med bolkene og deres antall henvendelser

henvser = ant_henv_per_bolk(kl_slett) # kjører funksjonen med arrayet inneholdende klokkeslett for henvendelsene som input/argument og tar i mot resultatet i 'henvser'

# oppretter kakediagram for å visualisere resultatet
labels = henvser.keys()
sizes = henvser.values()

plt.figure(figsize=(6, 6))
plt.pie(sizes, labels=labels, startangle=90)
plt.title("Henvendelser fordelt på totimersbolker")
plt.show()

# skriver ut resultatet med en løkke for alle items i henvser-dictionaryet
print('Antall henvendelser per totimersbolk\n')
for bolk, antall in henvser.items():
    print(f"{bolk}: {antall}")
