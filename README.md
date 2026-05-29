# 🦠 Uproszczony model symulacyjny epidemii

Rozwiązanie zadania **„Symulacja — zadanie 1”**: dyskretny, stochastyczny model
epidemii w stadzie **488 zwierząt**, obserwowany w *umownych jednostkach czasu*
(u.j.c.). Choroba ma charakter typu SIR z dwiema fazami zakaźnymi, a populacja
jest podzielona na **kohorty wiekowe** z reprodukcją i śmiertelnością zależną od
wieku.

Cała logika symulacji żyje w **formułach arkusza Excel** (`RAND`, `NORM.INV`,
`BINOM.INV`). Skrypty Pythona służą wyłącznie do *zbudowania* skoroszytu i
*weryfikacji* jego niezmienników — nie są środowiskiem uruchomieniowym modelu.

---

## 📂 Zawartość repozytorium

| Plik | Rola |
|------|------|
| `Symulacja zadanie 1.docx` | Treść zadania (parametry, reguły, termin). |
| `model epidemii (kohorty wiekowe).xlsx` | **Główny deliverable** — pełny model kohortowo-wiekowy (18 arkuszy, F9 = nowa trajektoria). |
| `uproszczony model epidemii.xlsx` | Wcześniejszy, zagregowany model (bez wymiaru wieku) — zachowany jako referencja. |
| `Uproszczony model symulacyjny epidemii.pptx` | Slajdy objaśniające dynamikę i rekurencję. |
| `build_model.py` | Jednorazowy builder skoroszytu (openpyxl). Zmiana parametru = edycja stałej na górze + ponowne uruchomienie. |
| `verify_model.py` | Re-implementacja logiki w Pythonie — sprawdza niezmienniki modelu. |
| `plan.txt` | Plan implementacji Zadania 1. |

---

## 🧬 Model choroby

Każde zwierzę przechodzi przez cykl:

```
        zakażenie                koniec fazy 2
  z  ──────────────▶  ca  ──▶  cb  ─────────────┬──▶  o  (odporność, 1 tick)  ──▶  z
(zdrowy           (faza 1:    (faza 2:          │
 podatny)        nosiciel     objawowy,         └──▶  u  (śmierć)
                 bezobjawowy)  zakaźny)
```

Każda faza trwa **dokładnie jeden tick**. Zarówno `ca`, jak i `cb` są zakaźne.

### Dynamika populacji

- **Reprodukcja** rusza od wieku 2. Współczynniki urodzeń (na tick, na liczebność kohorty):
  - wiek 2–4 → ~N(0.15, 0.02)
  - wiek 5+  → ~N(0.10, 0.01)
- **Transmisja pionowa**: P(noworodek chory) = (ca + cb) / N kohorty rodzicielskiej.
- **Śmiertelność chorobowa** (P(śmierć | w fazie `cb`)):
  - wiek 1–3 → ~N(0.20, 0.05)
  - wiek 4–5 → ~N(0.30, 0.07)
  - wiek 6+  → ~N(0.50, 0.15)
- **Śmiertelność naturalna**: warunkowy hazard `h_nat(a)` wyprowadzony z rozkładu długości życia N(6, 1), ograniczony do `h_nat(10) = 1`. Średnia długość życia ≈ 6 u.j.c.

### Stan początkowy (suma = 488)

| wiek |  z |  o | ca | cb |
|-----:|---:|---:|---:|---:|
| 1 | 60 | 20 | 10 | 10 |
| 2 | 60 | 30 | 10 | 20 |
| 3 | 70 | 10 |  5 | 10 |
| 4 | 60 | 10 | 10 |  5 |
| 5 | 20 | 20 |  7 |  3 |
| 6 | 10 |  5 |  6 |  4 |
| 7 | 10 |  0 |  0 |  3 |

---

## 📐 Rekurencja (na kohortę wiekową)

```
z(a, t+1)  = z(a−1, t) + o(a−1, t) − pca(a−1, t+1) − śmierci naturalne
ca(a, t+1) = pca(a−1, t+1)                          − śmierci naturalne
cb(a, t+1) = ca(a−1, t)                             − śmierci naturalne
o(a, t+1)  = cb(a−1, t) − pu(a−1, t+1)              − śmierci naturalne
u(t+1)     = u(t) + Σ pu + Σ śmierci naturalne
```

Noworodki wchodzą do kohorty `a = 1` w `t+1`, z podziałem na zdrowe / chore (faza 1).

---

## ▶️ Jak uruchomić

**Symulacja (Excel).** Otwórz `model epidemii (kohorty wiekowe).xlsx` i naciśnij
**F9**, aby wylosować nową trajektorię. Arkusz `Wyniki` zawiera agregaty per-tick,
wykres i statystyki podsumowujące. Arkusz `MonteCarlo` ma rusztowanie pod tablicę
danych do replikacji N przebiegów (konfiguracja ręczna).

**Przebudowa skoroszytu z parametrów:**

```bash
pip install openpyxl
python3 build_model.py        # generuje "model epidemii (kohorty wiekowe).xlsx"
```

**Weryfikacja niezmienników** (czysty Python, bez zależności):

```bash
python3 verify_model.py
```

Sprawdza, że: populacja t=0 = 488, wszystkie liczebności są nieujemnymi liczbami
całkowitymi, zachodzi `alive(t) + zgony_skum(t) − urodzenia_skum(t) = 488` w każdym
ticku oraz że oczekiwana długość życia ≈ 6.

### Przykładowy wynik weryfikacji

```
Initial total population: 488

--- Invariant checks ---
  population conservation: PASS
  non-negativity:          PASS

--- Monte Carlo sweep (500 reps) ---
           alive at t=30: mean     4.44  std    3.50  p5  0  p50  4  p95  10
      cum deaths at t=30: mean   926.64  std   42.11  p5 859  p50 925  p95 995
         peak infectious: mean   150.88  std   15.64  p5 126  p50 151  p95 176
  extinction rate: 8.8%

--- Disease-free lifespan sanity ---
  expected age at natural death from age 1: 6.50  (target ≈ 6)
```

---

## 🗂️ Struktura skoroszytu `model epidemii (kohorty wiekowe).xlsx`

| Arkusz | Zawartość |
|--------|-----------|
| `Założenia` | Jedyne źródło prawdy: tolerancje urodzeń/zgonów, tabela `h_nat(a)`, stan początkowy, kolejność operacji, zakresy nazwane. |
| `z`, `o`, `ca`, `cb` | Liczebności kompartmentów, 31 × 10 (tick × wiek). Wspólna konwencja wierszy/kolumn dla czystych sum międzyarkuszowych. |
| `pch`, `r_st`, `m_st`, `pca`, `pu`, `urodz`, `urodz_ch`, `nat_z/o/ca/cb` | Pośrednie losowania i liczniki per-tick. |
| `Wyniki` | Agregaty per-tick + wykres + statystyki + zakresy nazwane dla MC. |
| `MonteCarlo` | Rusztowanie tablicy danych (konfiguracja ręczna). |

> **Uwaga:** `row 2 = t=0`, `row 32 = t=30`; `kol B = wiek 1`, `kol K = wiek 10`.
> Arkusz nie używa odwołań strukturalnych — zmieniając formułę, propaguj edycję
> w dół całej kolumny.

---

## 🔤 Legenda oznaczeń

Konwencja `<stan>(t)` dla liczebności oraz `p<stan>` / `s<stan>` dla przyrostu /
spadku:

| Skrót | Znaczenie |
|-------|-----------|
| `z` / `pz` / `sz`   | zdrowy podatny — liczba / przyrost / spadek |
| `o` / `po` / `so`   | ozdrowieniec z czasową odpornością |
| `ca` / `pca` / `sca`| chory faza 1 (bezobjawowy nosiciel) |
| `cb` / `pcb` / `scb`| chory faza 2 (objawowy) |
| `u` / `pu`          | zgony (skumulowane / nowe) |
| `pch`               | prawdopodobieństwo zakażenia ≈ (ca+cb)/total |
| `h_nat(a)`          | warunkowy hazard śmierci naturalnej w wieku `a` |

---

## 📝 Uwagi

- Nazwy zmiennych (`ca`, `cb`, `pch`, …) celowo pozostają w formie polskiej — ich
  zmiana zdesynchronizowałaby treść zadania, slajdy i skoroszyt.
- `build_model.py` i `verify_model.py` współdzielą **te same stałe i kolejność
  operacji** — przejście weryfikatora daje pewność, że arkusz jest spójny
  strukturalnie.
