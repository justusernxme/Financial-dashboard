import openpyxl
import pandas as pd

# Comprehensive Financial Data (2015-2024) for AAPL and MSFT
REAL_DATA = {
    "AAPL": {
        "years": [
            "2015",
            "2016",
            "2017",
            "2018",
            "2019",
            "2020",
            "2021",
            "2022",
            "2023",
            "2024",
        ],
        "revenue": [
            233715e6,
            215639e6,
            229234e6,
            265595e6,
            260174e6,
            274515e6,
            365817e6,
            394328e6,
            383285e6,
            391035e6,
        ],
        "cogs": [
            140089e6,
            131376e6,
            141048e6,
            163756e6,
            161782e6,
            169559e6,
            212981e6,
            223546e6,
            214137e6,
            210352e6,
        ],
        "rd": [
            8067e6,
            10045e6,
            11581e6,
            14236e6,
            16217e6,
            18752e6,
            21914e6,
            26251e6,
            29915e6,
            31370e6,
        ],
        "sga": [
            14329e6,
            14194e6,
            15261e6,
            16705e6,
            18245e6,
            19916e6,
            21973e6,
            25094e6,
            24932e6,
            26097e6,
        ],
        "net_income": [
            53394e6,
            45687e6,
            48351e6,
            59531e6,
            55256e6,
            57411e6,
            94680e6,
            99803e6,
            96995e6,
            93736e6,
        ],
        "ebit": [
            72515e6,
            61372e6,
            61344e6,
            70898e6,
            63930e6,
            66288e6,
            108949e6,
            119437e6,
            114301e6,
            123216e6,
        ],
        "interest_expense": [
            1331e6,
            1456e6,
            2323e6,
            3240e6,
            3576e6,
            2873e6,
            2645e6,
            2931e6,
            3933e6,
            3576e6,
        ],
        "cash": [
            21120e6,
            20484e6,
            20289e6,
            25913e6,
            48844e6,
            38016e6,
            34940e6,
            23646e6,
            29965e6,
            29943e6,
        ],
        "mkt_sec": [
            20481e6,
            46671e6,
            53892e6,
            40388e6,
            51713e6,
            52927e6,
            27699e6,
            24658e6,
            31590e6,
            35228e6,
        ],
        "ar": [
            16849e6,
            15754e6,
            17874e6,
            23186e6,
            22926e6,
            16120e6,
            26278e6,
            28184e6,
            29508e6,
            33410e6,
        ],
        "inv": [
            2349e6,
            2132e6,
            4855e6,
            3956e6,
            4106e6,
            4061e6,
            6580e6,
            4946e6,
            6331e6,
            7286e6,
        ],
        "curr_assets": [
            89378e6,
            106869e6,
            128645e6,
            131339e6,
            162819e6,
            143713e6,
            134836e6,
            135405e6,
            143566e6,
            152987e6,
        ],
        "total_assets": [
            290479e6,
            321686e6,
            375319e6,
            365725e6,
            338516e6,
            323888e6,
            351002e6,
            352755e6,
            352583e6,
            364980e6,
        ],
        "curr_liab": [
            80610e6,
            79006e6,
            100814e6,
            115929e6,
            105718e6,
            105392e6,
            125481e6,
            153982e6,
            145308e6,
            176392e6,
        ],
        "total_debt": [
            64462e6,
            87032e6,
            115680e6,
            114483e6,
            108047e6,
            112436e6,
            124719e6,
            120069e6,
            111088e6,
            106629e6,
        ],
        "equity": [
            119355e6,
            128249e6,
            134047e6,
            107147e6,
            90488e6,
            65339e6,
            63090e6,
            50672e6,
            62146e6,
            56950e6,
        ],
        "ocf": [
            81266e6,
            65824e6,
            63598e6,
            77434e6,
            69391e6,
            80674e6,
            104038e6,
            122151e6,
            110543e6,
            118254e6,
        ],
    },
    "MSFT": {
        "years": [
            "2015",
            "2016",
            "2017",
            "2018",
            "2019",
            "2020",
            "2021",
            "2022",
            "2023",
            "2024",
        ],
        "revenue": [
            93580e6,
            85320e6,
            89950e6,
            110360e6,
            125843e6,
            143015e6,
            168088e6,
            198270e6,
            211915e6,
            245122e6,
        ],
        "cogs": [
            33038e6,
            32780e6,
            34261e6,
            38353e6,
            42910e6,
            46078e6,
            52232e6,
            62650e6,
            65863e6,
            74114e6,
        ],
        "rd": [
            12046e6,
            11988e6,
            13037e6,
            14726e6,
            16876e6,
            19269e6,
            20716e6,
            24512e6,
            27195e6,
            29510e6,
        ],
        "sga": [
            20324e6,
            19198e6,
            20116e6,
            22223e6,
            23098e6,
            24709e6,
            25224e6,
            27725e6,
            29589e6,
            32065e6,
        ],
        "net_income": [
            12193e6,
            16798e6,
            21204e6,
            16571e6,
            39240e6,
            44281e6,
            61271e6,
            72738e6,
            72361e6,
            88136e6,
        ],
        "ebit": [
            20165e6,
            22617e6,
            29331e6,
            35058e6,
            42959e6,
            52959e6,
            69916e6,
            83383e6,
            88523e6,
            102552e6,
        ],
        "interest_expense": [
            1378e6,
            121,
            2222e6,
            2733e6,
            2686e6,
            2591e6,
            2346e6,
            2063e6,
            1968e6,
            2163e6,
        ],
        "cash": [
            5595e6,
            6510e6,
            7663e6,
            11946e6,
            11356e6,
            13576e6,
            14224e6,
            13931e6,
            34704e6,
            18315e6,
        ],
        "mkt_sec": [
            90931e6,
            106730e6,
            125318e6,
            121822e6,
            122463e6,
            122951e6,
            116110e6,
            90826e6,
            76558e6,
            57228e6,
        ],
        "ar": [
            17908e6,
            18277e6,
            19792e6,
            26481e6,
            29524e6,
            32011e6,
            38043e6,
            44261e6,
            48688e6,
            56924e6,
        ],
        "inv": [
            2902e6,
            2251e6,
            2181e6,
            2662e6,
            2063e6,
            1895e6,
            2636e6,
            3742e6,
            2500e6,
            1246e6,
        ],
        "curr_assets": [
            124712e6,
            139660e6,
            159851e6,
            169662e6,
            175552e6,
            181915e6,
            184406e6,
            169684e6,
            184257e6,
            159734e6,
        ],
        "total_assets": [
            176223e6,
            193694e6,
            241086e6,
            258848e6,
            286556e6,
            301311e6,
            333779e6,
            364840e6,
            411976e6,
            512163e6,
        ],
        "curr_liab": [
            49858e6,
            59357e6,
            64527e6,
            58488e6,
            69420e6,
            72310e6,
            88657e6,
            95082e6,
            104149e6,
            125286e6,
        ],
        "total_debt": [
            35292e6,
            53687e6,
            86194e6,
            81808e6,
            78366e6,
            70998e6,
            67775e6,
            61270e6,
            59965e6,
            67127e6,
        ],
        "equity": [
            80083e6,
            71997e6,
            72394e6,
            82718e6,
            102330e6,
            118304e6,
            141988e6,
            166542e6,
            206223e6,
            268477e6,
        ],
        "ocf": [
            29080e6,
            33325e6,
            39507e6,
            43884e6,
            52185e6,
            60675e6,
            76740e6,
            89035e6,
            87582e6,
            118548e6,
        ],
    },
}


def update_ratios(ticker):
  file_name = f"{ticker}_Financial_Analysis.xlsx"
  d = REAL_DATA[ticker]

  try:
    wb = openpyxl.load_workbook(file_name)
  except Exception as e:
    print(f"Skipping {file_name}: {e}")
    return

  if "Financial Ratios" in wb.sheetnames:
    ws = wb["Financial Ratios"]
    for i in range(10):
      col = 3 + i  # Column C is 2015

      # Calculations
      curr = d["curr_assets"][i] / d["curr_liab"][i]
      quick = (d["cash"][i] + d["mkt_sec"][i] + d["ar"][i]) / d["curr_liab"][i]
      cash_r = (d["cash"][i] + d["mkt_sec"][i]) / d["curr_liab"][i]
      ocf_r = d["ocf"][i] / d["curr_liab"][i]

      dte = d["total_debt"][i] / d["equity"][i]
      dta = d["total_debt"][i] / d["total_assets"][i]
      lev = d["total_assets"][i] / d["equity"][i]
      interest_cov = (
          d["ebit"][i] / d["interest_expense"][i]
          if d["interest_expense"][i] > 0
          else 99.0
      )

      ar_turn = (
          d["revenue"][i] / d["ar"][i] if d["ar"][i] > 0 else 0
      )
      dso = 365 / ar_turn if ar_turn > 0 else 0
      inv_turn = (
          d["cogs"][i] / d["inv"][i] if d["inv"][i] > 0 else 0
      )
      dio = 365 / inv_turn if inv_turn > 0 else 0
      asset_turn = d["revenue"][i] / d["total_assets"][i]

      net_margin = d["net_income"][i] / d["revenue"][i]
      roe = d["net_income"][i] / d["equity"][i]
      roa = d["net_income"][i] / d["total_assets"][i]

      # Map to rows
      for row in range(1, ws.max_row + 1):
        label = str(ws.cell(row=row, column=1).value or "").lower()
        if "current ratio" in label:
          ws.cell(row=row, column=col, value=curr)
        elif "quick ratio" in label:
          ws.cell(row=row, column=col, value=quick)
        elif "cash ratio" in label:
          ws.cell(row=row, column=col, value=cash_r)
        elif "operating cash flow ratio" in label:
          ws.cell(row=row, column=col, value=ocf_r)
        elif "debt-to-equity" in label:
          ws.cell(row=row, column=col, value=dte)
        elif "debt-to-assets" in label:
          ws.cell(row=row, column=col, value=dta)
        elif "financial leverage" in label or "leverage ratio" in label:
          ws.cell(row=row, column=col, value=lev)
        elif "interest coverage" in label:
          ws.cell(row=row, column=col, value=interest_cov)
        elif "receivables turnover" in label:
          ws.cell(row=row, column=col, value=ar_turn)
        elif "days sales outstanding" in label or "dso" in label:
          ws.cell(row=row, column=col, value=dso)
        elif "inventory turnover" in label:
          ws.cell(row=row, column=col, value=inv_turn)
        elif "days inventory" in label or "dio" in label:
          ws.cell(row=row, column=col, value=dio)
        elif "asset turnover" in label:
          ws.cell(row=row, column=col, value=asset_turn)
        elif "net profit margin" in label or "net margin" in label:
          ws.cell(row=row, column=col, value=net_margin)
        elif "return on equity" in label or "roe" in label:
          ws.cell(row=row, column=col, value=roe)
        elif "return on assets" in label or "roa" in label:
          ws.cell(row=row, column=col, value=roa)

  wb.save(file_name)
  print(f"✅ Fully updated all ratios for {ticker}!")


if __name__ == "__main__":
  update_ratios("AAPL")
  update_ratios("MSFT")
