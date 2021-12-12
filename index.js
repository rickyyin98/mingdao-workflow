const input = process.argv[2];
const Excel = require("exceljs");
const path = require("path");
const Moment = require("moment");
const MomentRange = require("moment-range");
const moment = MomentRange.extendMoment(Moment);
const { XlsParser } = require("simple-excel-to-json");

var parser = new XlsParser();
(async () => {
  const workbook = new Excel.Workbook();
  const inputPath = path.normalize(input);
  const inputWorkbook = parser.parseXls2Json(inputPath, { isNested: true });
  let data = inputWorkbook[0].map((r) => Object.values(r));
  //   data = data.slice(1);
  const output = {
    price: {},
    total: {},
    count: {},
  };
  
  const startDate = moment("2020-07", "YYYY-MM").toDate().valueOf();
  const endDate = moment("2022-12", "YYYY-MM").toDate().valueOf();
  const dateIndex = 4;
  data = data.filter((row) => {
    const m = moment(row[dateIndex], "YYYY-MM-DD");
    const mv = m.toDate().valueOf();
    return mv >= startDate && mv <= endDate;
  });
  const groupedTotal = {};
  const clientMeta = {}
  data = data.map(
    ([client, industry, type, sku, date, monthes, count, price], i) => {
        clientMeta[client] = {
            industry, type, sku
        }
      const m = moment(date, "YYYY-MM-DD");
      const theMonth = m.format("YYYY-MM");
      if (!output["price"][i]) {
        output["price"][i] = {};
      }
      if (!output["count"][i]) {
        output["count"][i] = {};
      }
      if (!output["total"][i]) {
        output["total"][i] = {};
      }
      if (!output["price"][i][client]) {
        output["price"][i][client] = {};
      }
      if (!output["count"][i][client]) {
        output["count"][i][client] = {};
      }
      if (!output["total"][i][client]) {
        output["total"][i][client] = {};
      }
      // handle client monthes
      if (monthes > 1) {
        const startMonth = moment(theMonth, "YYYY-MM");
        const endMonth = moment(startMonth).add(
          monthes === 1 ? 1 : monthes - 1,
          "month"
        );
        const range = moment.range(startMonth, endMonth);
        // console.log(range)

        for (let date of range.by("month")) {
          const m = date.format("YYYY-MM");
          output["price"][i][client][m] = price;
          output["count"][i][client][m] = count;
          output["total"][i][client][m] = price * count;
          if(!groupedTotal[client]) {
            groupedTotal[client] = {}
          }
          groupedTotal[client][m] = price * count;
        }
      } else {
        output["count"][i][client][theMonth] = count;
        output["price"][i][client][theMonth] = price;
        output["total"][i][client][theMonth] = price * count;
        if(!groupedTotal[client]) {
            groupedTotal[client] = {}
          }
        groupedTotal[client][theMonth] = price * count;
      }

      return [
        client,
        industry,
        type,
        sku,
        m.toDate().valueOf(),
        monthes,
        count,
        price,
      ];
    }
  );

  const range = moment.range(startDate, endDate);
  // console.log(range)
  const monthes = [];
  for (let date of range.by("month")) {
    const theMonth = date.format("YYYY-MM");
    monthes.push(theMonth);
  }
  // console.log(monthes);

  data = data.map(
    ([client, industry, type, sku, date, monthes, count, price]) => {
      return [
        client,
        industry,
        type,
        sku,
        moment(date).format("YYYY-MM-DD"),
        monthes,
        count,
        price,
      ];
    }
  );
  // const datesArr = Array.from(dataDates).map( d => moment(d) ).sort( (a,b) => a.toDate() - b.toDate()).map( m => m.format('YYYY-MM-DD'));
  // console.log(datesArr)
  // console.log(data)
  // console.log(output)
  const priceSheet = workbook.addWorksheet("price");
  const countSheet = workbook.addWorksheet("count");
  const totalSheet = workbook.addWorksheet("total");
  const clientMRR = workbook.addWorksheet("客户维度MRR");
  priceSheet.addRow(["", "行业", "客户类型", "SKU", ...monthes]);
  countSheet.addRow(["", "行业", "客户类型", "SKU", ...monthes]);
  totalSheet.addRow(["", "行业", "客户类型", "SKU", ...monthes]);
  clientMRR.addRow(["","行业", "客户类型", "SKU",  ...monthes]);
  data.forEach(([client, industry, type, sku, _], i) => {
    priceSheet.addRow([
      client,
      industry,
      type,
      sku,
      ...monthes.map((d) => {
        if (!output["price"][i]) {
          return "";
        }
        if (!output["price"][i][client]) {
          return "";
        }
        return output["price"][i][client][d] || "";
      }),
    ]);
    countSheet.addRow([
      client,
      industry,
      type,
      sku,
      ...monthes.map((d) => {
        if (!output["count"][i]) {
          return "";
        }
        if (!output["count"][i][client]) {
          return "";
        }
        return output["count"][i][client][d] || "";
      }),
    ]);
    totalSheet.addRow([
      client,
      industry,
      type,
      sku,
      ...monthes.map((d) => {
        if (!output["total"][i]) {
          return "";
        }
        if (!output["total"][i][client]) {
          return "";
        }
        return output["total"][i][client][d] || "";
      }),
    ]);
  });

  Object.entries(groupedTotal).forEach(([key, values]) => {
    clientMRR.addRow([
        key,
        clientMeta[key].industry,
        clientMeta[key].type,
        clientMeta[key].sku,
        ...monthes.map((m) => {
          if (!groupedTotal[key]) {
            return "";
          }
          if (!groupedTotal[key][m]) {
            return "";
          }

          return groupedTotal[key][m] || "";
        }),
      ]);

  });

  await workbook.xlsx.writeFile(`output.xlsx`);
})();
