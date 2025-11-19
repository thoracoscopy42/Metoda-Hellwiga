# Metoda Hellwiga
Metoda Hellwiga służy do oceny pojemności informacyjnej zestawów zmiennych objaśniających, uwzględniając zarówno ich związek ze zmienną zależną, jak i współzależności pomiędzy predyktorami tak by wnosiły możliwie największą informację statystyczną.

Dla danej konfiguracji zmiennych $$L_k\$$ **indywidualna pojemność informacyjna** zmiennej $$X_j\$$ określana jest zależnością:

$$h_{kj} = \frac{r_{j}^{2}}{\sum_{i \in L_k} \lvert r_{ij} \rvert}$$

Całkowitą pojemność informacyjną kombinacji, czyli **pojemnością integralną kombinacji** stanowi suma indywidualnych pojemności nośników informacji $$h_{kj}\$$ należących do danej konfiguracji:

$$H_k = \sum_{j \in L_k} h_{kj}$$

Celem programu jest ułatwienie studentom i badaczom wyboru optymalnych zmiennych objaśniających do modeli statystycznych poprzez automatyzację tejże metody.

# Instalacja
1. Pobierz najnowszą wersję programu z zakładki <b>Releases</b>, w sekcji **Assets** znajduje się plik wykonywalny <i>Hellwig.exe.</i>.

   <img width="311" height="152" alt="image" src="https://github.com/user-attachments/assets/67489025-5219-4014-b0bc-f828a14d329e" />
   
3. Umieść plik <i>Hellwig.exe.</i> w dowolnym folderze.

Do użytkowania programu nie trzeba, poza samym plikiem <b><i>Hellwig.exe</i></b> nic pobierać, ani instalować - jest gotowy do użycia od razu.

<details>
<summary><strong> (Opcjonalnie) uruchomienie wersji źródłowej</strong></summary>

Jeśli chcesz uruchomić program ze środowiskiem programistycznym, wymagane pakiety znajdują się w pliku `requirements.txt`.

Możesz je zainstalować poleceniem:
```bash
pip install -r requirements.txt
```

</details>

# Jak używać programu
Najlepiej przygotować plik Excel zawierający jeden arkusz z macierzą korelacji.

1. Przygotuj w Excelu macierz, która jest w formacie:
<img width="561" height="206" alt="image" src="https://github.com/user-attachments/assets/a027cd83-25fb-4010-9b44-bf4e47feebb6" />


Musi ona zawierać jedną kolumnę z nazwą `Y` - będzie to zmienna objaśniana.

2. Uruchom <b><i>Hellwig.exe.</i></b>

3. Upewnij się, że excel w którym znajduje się macierz korelacji jest zamknięty przed uruchomieniem obliczeń.

4. Wybierz plik excel i kliknij <i>Uruchom Metodę Hellwiga.</i>

5. W oknie aplikacji pojawi się 10 najlepszych rezultatów, a w nowym arkuszu zostaną dodane wszystkie obliczone kombinacje.




