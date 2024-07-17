1. CLone this Repo.

2. Setup Virtual Environment

    python -m venv venv

3. Activate Virtual Environment & install requirements

    venv\Script\activate

    pip install -r requirements.txt

5. Load the '.hse' files in "HSE Files" manually by copying '.hse' files from "DELFOR" directory and paste here in this project directory.

6. Run the 'WorkBookandPDFGenerator.py' file to generate the main workbook and pdfs with sales orders, picklists, and forecasts

    python WorkBookandPDFGenerator.py
   
``Enter a date in 'yyyy, m, d' format or type 'p' to use the current date:``
   --press 'p' to have the current week as 'first week'

   ~ p  {if this week (current week) is the first week}

   ~ 2022, 8, 19 {if 19/08/2022 -- is the first week}

7. If copy & pasted HSE files also includes 'WPCX' -- Services file, the press the following command.

  ``Also Prepare the Services Picklist, Type Y for Yes and N for No``
  --  y {to also create Services Picklist as well in .Report\picklist\services}
  -- n {if not required.}

8. Forecasts Directory

example - Weekly Forecasts -- C:\Users\Rama\JustInTime\Report\SalesForecast\Weekly_Forecast_17_07_2024_1500.pdf
example - Monthly Forecasts -- C:\Users\Rama\JustInTime\Report\SalesForecast\Monthly_Forecast_17_07_2024_1500.pdf
  
9. Picklists Directory

   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM002_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM004_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM005_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM007_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM008_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM009_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM0011_Picklist_17_07_2024_1500.pdf
   example - Sepearate Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Separate_sheets\BAM0018_Picklist_17_07_2024_1500.pdf

   example - Full Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\FullPicklist\Picklist17_07_2024_1500.pdf

   example - Services Picklist -- C:\Users\Rama\JustInTime\Report\Picklist\Services\BAM003_Services_Picklist_17_07_2024_1506.pdf
   


   
