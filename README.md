# Excel kolonnu kārtotājs
## Projekta uzdevums
Projekta uzdevums ir sakārtot lietotāja izvēlētu Excel faila kolonnu alfabētiski, skaitliski vai pēc datuma, augošā vai dilstošā secībā, atkarībā no lietotāja prasītā un saglabāt jaunā lapā. Kļūdaini uzrakstītas vērtības tiek novietotas kolonnas sākumā vai beigās, atkarībā no tā vai tiek kārtots augošā vai dilstošā secībā. 

> [!WARNING]
> Kļūdaini uzrakstītas vērtības skaitās kā simboli un cipari kolonnā, kura tiek kārtota alfabētiski; burti un simboli kolonnā, kura tiek kārtota skaitliski vai pēc datuma.  

Projektam ir pievienots piemēra Excel fails (ex.xlsx), kuru var izmantot, lai izmēģinātu programmu.
## Izmantotās bibliotēkas
### Izmantotas sys un pandas bibliotēkas
**sys** bibliotēka tika izmantota, lai izietu no programmas, ja lietotājs ievadījis nepareizus vai kļūdainus datus.  
**pandas** bibliotēka tika izmantota datu apstrādei un Excel faila lasīšanai un papildināšanai.
## Programmatūras izmantošanas metode
1. Izpilda skriptu un ievada Excel faila pilno nosaukumu, ievērojot simbolus, lielos un mazos burtus.
2. Ievada faila lapas nosaukumu, ievērojot simbolus, lielos un mazos burtus.
3. Ievada, pēc kādiem parametriem vēlas sakārtot kolonnu, izvēloties vienu no piedāvātajām opcijām, ievērojot galotni. Netiek ievēroti lielie un mazie burti, garumzīmes.
4. Ievada, vai datus kārtot augošā vai dilstošā secībā, ievērojot galotni. Netiek ievēroti lielie un mazie burti, garumzīmes.
5. Ievada tās kolonnas, kuru vēlas sakārtot, nosaukumu. Kolonnas nosaukumam jābūt uzrakstītam tā pat, kā Excel failā.
6. Ievada nosaukumu Excel lapai, kurā sakārtotie dati tiks saglabāti.