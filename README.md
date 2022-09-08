# VBA_Analysis

The following analysis was conducted for Steve to consider total daily volume and return for several different stocks. Naturally, the data included for 2017 and 2018 contains entries for every day of the year for each stock. Therefore, to consolidate the data to make conclusions, VBA functionality was utilized. 

![Total_Daily_Volume_2017](https://user-images.githubusercontent.com/111781762/189117722-fdf6bb70-84d2-47ff-9fdf-4fb4a9966c50.png)

![Total_Daily_Volume_2018](https://user-images.githubusercontent.com/111781762/189117733-d901b76b-4cfa-441b-a4e8-e2417d6b90bd.png)

As we can observe, Daily Volume decreased in 2018 for tickers AY, CSIQ, RUN, SEDG, TERP and VSLR. For a majority of stocks that exhibited higher trading volume, return did not meet projections and return decreased. Therefore, the following stocks chosen by Steve may not represent great potential investments for his clients. 

![Return_For_2017](https://user-images.githubusercontent.com/111781762/189119927-bc7e0cff-4046-4fee-acbe-b7df0c0026dd.png)

![Return_For_2018](https://user-images.githubusercontent.com/111781762/189119945-19ff7d78-a965-48b8-a3a3-b39bda4d9970.png)

As we can see, return plumetted for all stocks included in the list that do not include tickers ENPH and RUN. Therefore, based on past performance, ENPH and RUN may be the most viable options for investment. Further analysis must be conducted on regarding internal operations and structure for each to justify investment.

![VBA Run Time](https://user-images.githubusercontent.com/111781762/189120897-6336ce48-84d1-49b1-ad09-ba7ba097e888.png)

Originally, VBA was utilized in a less timely manner that involved unnecessary usage of loops. This decreased ability to process data considerably. Therefore, the VBA module was rewritten to increase performance. This should help considerably should a larger set of data with more stocks be used in the future. The following VBA script can be changed for this purpose.
