#pragma execution_character_set("utf-8")

#include "dataInputXLSX.h"


fileXLSX::fileXLSX(const std::string& fileName_, const std::string& sheetName_, const std::string& tableNameCreated_)
{
	this->fileName = fileName_;
	this->sheetName = sheetName_;
	this->tableNameCreated = tableNameCreated_;
	
	try {
		openDoc.open(fileName);

		auto wks = openDoc.workbook().worksheet(sheetName);
		
		int priceColNum = getColumnNumber(sharePriceNameStr);
		int nameCodeColNum = getColumnNumber(shareCodeNameStr);
		int lotColNum = getColumnNumber(shareLotNameStr);
		int priceMoveColNum = getColumnNumber(priceMoveNameStr);
		int bestSupplyColNum = getColumnNumber(bestSupplyNameStr);
		int bestDemandColNum = getColumnNumber(bestDemandNameStr);
		int minPriceColNum = getColumnNumber(minPriceStr);
		int maxPriceColNum = getColumnNumber(maxPriceStr);
		int assetVolumeColNum = getColumnNumber(assetVolumeStr);
		int openPriceColNum = getColumnNumber(openPriceStr);

		
		int numberRows = static_cast<int>(wks.rowCount());

		for (int row = 2; row <= numberRows; ++row)
		{
			// Первый аргумент - номер строки, второй - номер столбца
			OpenXLSX::XLCellValue cell = wks.cell(row, priceColNum).value();
			double price = 0.f;			
			if (cell.type() == OpenXLSX::XLValueType::Float || cell.type() == OpenXLSX::XLValueType::Integer)
			{
				std::string stockName = wks.cell(row, nameCodeColNum).value().get<std::string>();
				int lot = wks.cell(row, lotColNum).value().get<int>();
				if (cell.type() == OpenXLSX::XLValueType::Float)
				{
					price = cell.get<double>();
				}
				else if (cell.type() == OpenXLSX::XLValueType::Integer)
				{
					price = static_cast<double>(cell.get<int>());
				}

				double priceMove = 0;
				OpenXLSX::XLCellValue priceMoveCell = wks.cell(row, priceMoveColNum).value();
				if (priceMoveCell.type() == OpenXLSX::XLValueType::Float)
				{
					priceMove = priceMoveCell.get<double>();
				}
				else if (priceMoveCell.type() == OpenXLSX::XLValueType::Integer)
				{
					priceMove = static_cast<double>(priceMoveCell.get<int>());
				}

				double supply = 0;
				OpenXLSX::XLCellValue supplyCell = wks.cell(row, bestSupplyColNum).value();
				if (supplyCell.type() == OpenXLSX::XLValueType::Float)
				{
					supply = supplyCell.get<double>();
				}
				else if (supplyCell.type() == OpenXLSX::XLValueType::Integer)
				{
					supply = static_cast<double>(supplyCell.get<int>());
				}

				double demand = 0;
				OpenXLSX::XLCellValue demandCell = wks.cell(row, bestDemandColNum).value();
				if (demandCell.type() == OpenXLSX::XLValueType::Float)
				{
					demand = demandCell.get<double>();
				}
				else if (demandCell.type() == OpenXLSX::XLValueType::Integer)
				{
					demand = static_cast<double>(demandCell.get<int>());
				}

				double minP = 0;
				OpenXLSX::XLCellValue minCell = wks.cell(row, minPriceColNum).value();
				if (minCell.type() == OpenXLSX::XLValueType::Float)
				{
					minP = minCell.get<double>();
				}
				else if (minCell.type() == OpenXLSX::XLValueType::Integer)
				{
					minP = static_cast<double>(minCell.get<int>());
				}

				double maxP = 0;
				OpenXLSX::XLCellValue maxCell = wks.cell(row, maxPriceColNum).value();
				if (maxCell.type() == OpenXLSX::XLValueType::Float)
				{
					maxP = maxCell.get<double>();
				}
				else if (maxCell.type() == OpenXLSX::XLValueType::Integer)
				{
					maxP = static_cast<double>(maxCell.get<int>());
				}

				double assetVolume = 0;
				OpenXLSX::XLCellValue assetVolumeCell = wks.cell(row, assetVolumeColNum).value();
				if (assetVolumeCell.type() == OpenXLSX::XLValueType::Float)
				{
					assetVolume = assetVolumeCell.get<double>();
				}
				else if (assetVolumeCell.type() == OpenXLSX::XLValueType::Integer)
				{
					assetVolume = static_cast<double>(assetVolumeCell.get<int>());
				}

				double openPrice = 0;
				OpenXLSX::XLCellValue openPriceCell = wks.cell(row, openPriceColNum).value();
				if (openPriceCell.type() == OpenXLSX::XLValueType::Float)
				{
					openPrice = openPriceCell.get<double>();
				}
				else if (openPriceCell.type() == OpenXLSX::XLValueType::Integer)
				{
					openPrice = static_cast<double>(openPriceCell.get<int>());
				}


				stockVec.push_back(StockDataClass(stockName, price, lot, 
					priceMove, supply, demand, minP, maxP, assetVolume, openPrice));

			} else if (cell.type() == OpenXLSX::XLValueType::Empty)
			{
				std::cout << "Номер строки: " << row << " -> " <<
						"empty" << " " << std::endl;
			}
		}
		
		std::cout << std::endl;
		
	}
	catch (std::exception& e) {
		std::string catchedErrorName = e.what();
		std::cout << "Error open xlsx file: " << fileName << " " << catchedErrorName << std::endl;
		if (catchedErrorName == "file open failed")
		{
			std::cout << "Possible error: " << fileName << " file not found." << std::endl;
			std::cout << "Please ensure the file is in the correct directory and the file name is spelled correctly." << std::endl;
		}
		else if (catchedErrorName == "Cell reference is invalid")
		{
			std::cout << "Potential error : Required columns are absent from the file." << std::endl;
			std::cout << "Please verify that the input table includes all columns needed for processing." << std::endl;
		}

	}

}


fileXLSX::~fileXLSX()
{
	openDoc.save();
	openDoc.close();
	
}

std::vector<StockDataClass> fileXLSX::getAllDataFromFile()
{
	return std::vector<StockDataClass>();
}

void fileXLSX::printStartData()
{
	for (const auto& share : stockVec)
	{
		share.printStockData();
	}
}

void fileXLSX::printCalculatedData(double number, double riskProc)
{
	calData(number, riskProc);

	for (const auto& share : stockCalculatedVec)
	{
		share.printStockCalculatedData();
	}
}

void fileXLSX::printShareCalData(const std::string& comName, double amount, double riskProc)
{
	for (const auto& share : stockVec)
	{
		if (share.getCompanyName() == comName)
		{
			double lotPrice = share.getLotPrice();
			double lotMove = share.getLotMove();
			int lotShare = share.getInOneLotShares();
			double sharePrice = share.getSharePrice();
			double priceMove = share.getPriceMove();

			printShare(StockCalculateDataClass(comName, amount, lotShare, lotPrice, riskProc,
				lotMove, sharePrice, priceMove));
		}
	}
}

void fileXLSX::printDealData(StockDataOptDialStopProfitClass data)
{
	data.printStockDataOptDialStopProfit();
}

bool fileXLSX::createCalculatedTableFromTableData(double amount, double riskProc)
{
	//Стиль шрифта 
	auto bold_font = xlnt::font();
	bold_font.bold(true);
	bold_font.size(16);
	//Стиль границы
	xlnt::border border;
	xlnt::border::border_property border_property;
	border_property.style(xlnt::border_style::thin);
	border.side(xlnt::border_side::start, border_property); // left
	border.side(xlnt::border_side::end, border_property); // right
	border.side(xlnt::border_side::top, border_property); // top
	border.side(xlnt::border_side::bottom, border_property); // bottom
	//Выравнивание
	xlnt::alignment center;
	center.vertical(xlnt::vertical_alignment::center);
	center.horizontal(xlnt::horizontal_alignment::center);
	center.wrap(true);
	//Стиль ячеек под названием DealNames
	auto Deal_style = createdDoc.create_style("DealNames");
	Deal_style.font(bold_font);
	Deal_style.alignment(center);
	Deal_style.fill(xlnt::fill::solid(red_color));
	Deal_style.border(border);

	//Стиль ячеек под названием CalcNames
	auto calc_font = xlnt::font();
	calc_font.bold(true);
	calc_font.color(white_color);
	calc_font.size(12);

	auto Calc_style = createdDoc.create_style("CalcNames");
	Calc_style.font(calc_font);
	Calc_style.fill(xlnt::fill::solid(perpule_color));
	Calc_style.border(border);
	Calc_style.alignment(center);

	if (stockDataOptDialProfitVec.size() != 0)
	{
		auto newSheet = createdDoc.active_sheet();
		newSheet.title("ADFT"); //Чистая дата из таблицы, можно менять самому стоп и профит
		auto secSheet = createdDoc.create_sheet();
		secSheet.title("SPF");//Дата из таблицы, но уже формула на стоп и профит
		auto thirdSheet = createdDoc.create_sheet();
		thirdSheet.title("ProcDeals");//Расчет одного ассета, где стоп и профит расчитываются исходя из процентов
		auto fourthSheet = createdDoc.create_sheet();
		fourthSheet.title("Deals");//Расчет одного ассета, где можно ставить свои стоп и профит
		auto fifthSheet = createdDoc.create_sheet();
		fifthSheet.title("ForcastingALL");//Расчетные значения всех эмитентов на следующий день, исходя из текущих цен
		auto sixSheet = createdDoc.create_sheet();
		sixSheet.title("ForcastingOne");//Расчетные значения на одного эмитента на следующий день, исходя из текущих цен
		
		//Первая книга
		newSheet.cell("A1").value("Сумма для инвестиции");
		newSheet.cell("B1").value("Процент риска");
		newSheet.cell("C1").value("Сумма риска");

		auto deal_range1 = newSheet.range("A1:C1");
		deal_range1.style(Deal_style);

		auto cellАmount1 = newSheet.cell("A2");
		cellАmount1.number_format(xlnt::number_format::number_format("#,##₽"));
		cellАmount1.value(amount);
		makeNoFormulaCellStyle(cellАmount1);

		auto cellRiskProc1 = newSheet.cell("B2");
		cellRiskProc1.number_format(xlnt::number_format::percentage_00());
		cellRiskProc1.value(riskProc / 100);
		makeNoFormulaCellStyle(cellRiskProc1);

		auto cellRiskSum1 = newSheet.cell("C2");
		cellRiskSum1.number_format(xlnt::number_format::number_format("#,##₽"));
		cellRiskSum1.formula("=A2 * B2");
		makeCellStyle(cellRiskSum1);

		newSheet.cell("A4").value("Код инструмента");
		newSheet.cell("B4").value("Лот");
		newSheet.cell("C4").value("Т.Цена");
		newSheet.cell("D4").value("Шаг Цены");
		newSheet.cell("E4").value("Стоп Цена");
		newSheet.cell("F4").value("Профит Цена");
		newSheet.cell("G4").value("Цена лота");
		newSheet.cell("H4").value("MaxLotFromInvSum");
		newSheet.cell("I4").value("Макс Денежный Риск");
		newSheet.cell("J4").value("Макс Исп. Денг.");
		newSheet.cell("K4").value("Остаток Денег");
		newSheet.cell("L4").value("Макс Профит");
		newSheet.cell("M4").value("Коэф. Риск/Профит");

		auto calc_range1 = newSheet.range("A4:M4");
		calc_range1.style(Calc_style);
		calc_range1.alignment(center);

		//Вторая книга
		secSheet.cell("A1").value("Сумма для инвестиции");
		secSheet.cell("B1").value("Процент риска");
		secSheet.cell("C1").value("Сумма риска");
		secSheet.cell("D1").value("СтопЦена %от Цены");
		secSheet.cell("E1").value("Профит %от Цены");

		auto deal_range2 = secSheet.range("A1:E1");
		deal_range2.style(Deal_style);

		auto cellАmount2 = secSheet.cell("A2");
		cellАmount2.number_format(xlnt::number_format::number_format("#,##₽"));
		cellАmount2.value(amount);
		makeNoFormulaCellStyle(cellАmount2);

		auto cellRiskProc2 = secSheet.cell("B2");
		cellRiskProc2.number_format(xlnt::number_format::percentage_00());
		cellRiskProc2.value(riskProc / 100);
		makeNoFormulaCellStyle(cellRiskProc2);

		auto cellRiskSum2 = secSheet.cell("C2");
		cellRiskSum2.number_format(xlnt::number_format::number_format("#,##₽"));
		cellRiskSum2.formula("=A2 * B2");
		makeCellStyle(cellRiskSum2);

		auto cellStopProc2 = secSheet.cell("D2");
		cellStopProc2.number_format(xlnt::number_format::percentage_00());
		cellStopProc2.value(0.1);
		makeNoFormulaCellStyle(cellStopProc2);

		auto cellProfitProc2 = secSheet.cell("E2");
		cellProfitProc2.number_format(xlnt::number_format::percentage_00());
		cellProfitProc2.value(0.2);
		makeNoFormulaCellStyle(cellProfitProc2);

		secSheet.cell("A4").value("Код инструмента");
		secSheet.cell("B4").value("Лот");
		secSheet.cell("C4").value("Т.Цена");
		secSheet.cell("D4").value("Шаг Цены");
		secSheet.cell("E4").value("Стоп Цена");
		secSheet.cell("F4").value("Профит Цена");
		secSheet.cell("G4").value("Цена лота");
		secSheet.cell("H4").value("MaxLotFromInvSum");
		secSheet.cell("I4").value("Макс Денежный Риск");
		secSheet.cell("J4").value("Макс Исп. Денг.");
		secSheet.cell("K4").value("Остаток Денег");
		secSheet.cell("L4").value("Макс Профит");
		secSheet.cell("M4").value("Коэф. Риск/Профит");

		auto calc_range2 = secSheet.range("A4:M4");
		calc_range2.style(Calc_style);
		calc_range2.alignment(center);

		//Третья книга
		thirdSheet.cell("A1").value("Сумма для инвестиции");
		thirdSheet.cell("B1").value("Процент риска");
		thirdSheet.cell("C1").value("Сумма риска");
		thirdSheet.cell("D1").value("СтопЦена %от Цены");
		thirdSheet.cell("E1").value("Профит %от Цены");

		auto deal_range3 = thirdSheet.range("A1:E1");
		deal_range3.style(Deal_style);

		auto cellАmount3 = thirdSheet.cell("A2");
		cellАmount3.number_format(xlnt::number_format::number_format("#,##₽"));
		cellАmount3.value(amount);
		makeNoFormulaCellStyle(cellАmount3);

		auto cellRiskProc3 = thirdSheet.cell("B2");
		cellRiskProc3.number_format(xlnt::number_format::percentage_00());
		cellRiskProc3.value(riskProc / 100);
		makeNoFormulaCellStyle(cellRiskProc3);

		auto cellRiskSum3 = thirdSheet.cell("C2");
		cellRiskSum3.number_format(xlnt::number_format::number_format("#,##₽"));
		cellRiskSum3.formula("=A2 * B2");
		makeCellStyle(cellRiskSum3);

		auto cellStopProc3 = thirdSheet.cell("D2");
		cellStopProc3.number_format(xlnt::number_format::percentage_00());
		cellStopProc3.value(0.1);
		makeNoFormulaCellStyle(cellStopProc3);

		auto cellProfitProc3 = thirdSheet.cell("E2");
		cellProfitProc3.number_format(xlnt::number_format::percentage_00());
		cellProfitProc3.value(0.2);
		makeNoFormulaCellStyle(cellProfitProc3);

		thirdSheet.cell("A4").value("Код инструмента");
		thirdSheet.cell("B4").value("Лот");
		thirdSheet.cell("C4").value("Т.Цена");
		thirdSheet.cell("D4").value("Шаг Цены");
		thirdSheet.cell("E4").value("Стоп Цена");
		thirdSheet.cell("F4").value("Профит Цена");
		thirdSheet.cell("G4").value("Цена лота");
		thirdSheet.cell("H4").value("MaxLotFromInvSum");
		thirdSheet.cell("I4").value("Макс Денежный Риск");
		thirdSheet.cell("J4").value("Макс Исп. Денг.");
		thirdSheet.cell("K4").value("Остаток Денег");
		thirdSheet.cell("L4").value("Макс Профит");
		thirdSheet.cell("M4").value("Коэф. Риск/Профит");

		auto calc_range3 = thirdSheet.range("A4:M4");
		calc_range3.style(Calc_style);
		calc_range3.alignment(center);

		auto nameCell3 = thirdSheet.cell("A5");
		makeNoFormulaCellStyle(nameCell3);

		auto lotCell3 = thirdSheet.cell("B5");
		makeNoFormulaCellStyle(lotCell3);

		auto priceCell3 = thirdSheet.cell("C5");
		makeNoFormulaCellStyle(priceCell3);

		auto priceMoveCell3 = thirdSheet.cell("D5");
		makeNoFormulaCellStyle(priceMoveCell3);

		auto stopPriceCell3 = thirdSheet.cell("E5");
		stopPriceCell3.formula("= C5 - C5 * $D$2");
		makeCellStyle(stopPriceCell3);

		auto profitPriceCell3 = thirdSheet.cell("F5");
		profitPriceCell3.formula("= C5 + C5 * $E$2");
		makeCellStyle(profitPriceCell3);

		auto oneLotPriceCell3 = thirdSheet.cell("G5");
		oneLotPriceCell3.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		oneLotPriceCell3.formula("= C5 * B5");
		makeCellStyle(oneLotPriceCell3);

		auto maxLotToBuyCell3 = thirdSheet.cell("H5");
		maxLotToBuyCell3.formula("=IF($A$2 < C5 * B5 * "
			"ROUNDDOWN(($C$2/((C5 - E5)/ D5))"
			"/(B5 * D5),0), "
			"ROUNDDOWN($A$2/(C5 * B5),0), "
			"ROUNDDOWN(($C$2/((C5 - E5)/D5))"
			"/(B5 * D5),0))");
		makeCellStyle(maxLotToBuyCell3);

		auto maxMoneyRistCell3 = thirdSheet.cell("I5");
		maxMoneyRistCell3.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxMoneyRistCell3.formula("=H5 * B5 * "
			"D5 * ((C5 - E5)/D5) ");
		makeCellStyle(maxMoneyRistCell3);

		auto maxUseOfMoneyCell3 = thirdSheet.cell("J5");
		maxUseOfMoneyCell3.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxUseOfMoneyCell3.formula("=C5 * B5 * H5");
		makeCellStyle(maxUseOfMoneyCell3);

		auto restMoneyCell3 = thirdSheet.cell("K5");
		restMoneyCell3.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		restMoneyCell3.formula("=$A$2 - J5");
		makeCellStyle(restMoneyCell3);

		auto maxExpectedProfitCell3 = thirdSheet.cell("L5");
		maxExpectedProfitCell3.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxExpectedProfitCell3.formula("=H5 * D5 * B5 * "
			"((F5 - C5)/D5)");
		makeCellStyle(maxExpectedProfitCell3);

		auto riskProfitRatioCell3 = thirdSheet.cell("M5");
		riskProfitRatioCell3.number_format(xlnt::number_format::number_format("#,##0.000"));
		riskProfitRatioCell3.formula("=I5 / L5");
		makeCellStyle(riskProfitRatioCell3);

		//Четвертая книга
		fourthSheet.cell("A1").value("Сумма для инвестиции");
		fourthSheet.cell("B1").value("Процент риска");
		fourthSheet.cell("C1").value("Сумма риска");

		auto deal_range4 = fourthSheet.range("A1:C1");
		deal_range4.style(Deal_style);

		auto cellАmount4 = fourthSheet.cell("A2");
		cellАmount4.number_format(xlnt::number_format::number_format("#,##₽"));
		cellАmount4.value(amount);
		makeNoFormulaCellStyle(cellАmount4);

		auto cellRiskProc4 = fourthSheet.cell("B2");
		cellRiskProc4.number_format(xlnt::number_format::percentage_00());
		cellRiskProc4.value(riskProc / 100);
		makeNoFormulaCellStyle(cellRiskProc4);

		auto cellRiskSum4 = fourthSheet.cell("C2");
		cellRiskSum4.number_format(xlnt::number_format::number_format("#,##₽"));
		cellRiskSum4.formula("=A2 * B2");
		makeCellStyle(cellRiskSum4);

		fourthSheet.cell("A4").value("Код инструмента");
		fourthSheet.cell("B4").value("Лот");
		fourthSheet.cell("C4").value("Т.Цена");
		fourthSheet.cell("D4").value("Шаг Цены");
		fourthSheet.cell("E4").value("Стоп Цена");
		fourthSheet.cell("F4").value("Профит Цена");
		fourthSheet.cell("G4").value("Цена лота");
		fourthSheet.cell("H4").value("MaxLotFromInvSum");
		fourthSheet.cell("I4").value("Макс Денежный Риск");
		fourthSheet.cell("J4").value("Макс Исп. Денг.");
		fourthSheet.cell("K4").value("Остаток Денег");
		fourthSheet.cell("L4").value("Макс Профит");
		fourthSheet.cell("M4").value("Коэф. Риск/Профит");

		auto calc_range4 = fourthSheet.range("A4:M4");
		calc_range4.style(Calc_style);
		calc_range4.alignment(center);

		auto nameCell4 = fourthSheet.cell("A5");
		makeNoFormulaCellStyle(nameCell4);

		auto lotCell4 = fourthSheet.cell("B5");
		makeNoFormulaCellStyle(lotCell4);

		auto priceCell4 = fourthSheet.cell("C5");
		makeNoFormulaCellStyle(priceCell4);

		auto priceMoveCell4 = fourthSheet.cell("D5");
		makeNoFormulaCellStyle(priceMoveCell4);

		auto stopPriceCell4 = fourthSheet.cell("E5");
		makeNoFormulaCellStyle(stopPriceCell4);

		auto profitPriceCell4 = fourthSheet.cell("F5");
		makeNoFormulaCellStyle(profitPriceCell4);

		auto oneLotPriceCell4 = fourthSheet.cell("G5");
		oneLotPriceCell4.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		oneLotPriceCell4.formula("= C5 * B5");
		makeCellStyle(oneLotPriceCell4);

		auto maxLotToBuyCell4 = fourthSheet.cell("H5");
		maxLotToBuyCell4.formula("=IF($A$2 < C5 * B5 * "
			"ROUNDDOWN(($C$2/((C5 - E5)/ D5))"
			"/(B5 * D5),0), "
			"ROUNDDOWN($A$2/(C5 * B5),0), "
			"ROUNDDOWN(($C$2/((C5 - E5)/D5))"
			"/(B5 * D5),0))");
		makeCellStyle(maxLotToBuyCell4);

		auto maxMoneyRistCell4 = fourthSheet.cell("I5");
		maxMoneyRistCell4.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxMoneyRistCell4.formula("=H5 * B5 * "
			"D5 * ((C5 - E5)/D5) ");
		makeCellStyle(maxMoneyRistCell4);

		auto maxUseOfMoneyCell4 = fourthSheet.cell("J5");
		maxUseOfMoneyCell4.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxUseOfMoneyCell4.formula("=C5 * B5 * H5");
		makeCellStyle(maxUseOfMoneyCell4);

		auto restMoneyCell4 = fourthSheet.cell("K5");
		restMoneyCell4.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		restMoneyCell4.formula("=$A$2 - J5");
		makeCellStyle(restMoneyCell4);

		auto maxExpectedProfitCell4 = fourthSheet.cell("L5");
		maxExpectedProfitCell4.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxExpectedProfitCell4.formula("=H5 * D5 * B5 * "
			"((F5 - C5)/D5)");
		makeCellStyle(maxExpectedProfitCell4);

		auto riskProfitRatioCell4 = fourthSheet.cell("M5");
		riskProfitRatioCell4.number_format(xlnt::number_format::number_format("#,##0.000"));
		riskProfitRatioCell4.formula("=I5 / L5");
		makeCellStyle(riskProfitRatioCell4);


		//Пятая книга///////////////////////////////////////////

		fifthSheet.cell("A1").value("Код инструмента");
		fifthSheet.cell("B1").value("Т.Цена");
		fifthSheet.cell("C1").value("Цена Открытия");
		fifthSheet.cell("D1").value("Макс");
		fifthSheet.cell("E1").value("Мин");
		fifthSheet.cell("F1").value("коэф Х");
		fifthSheet.cell("G1").value("Х");
		fifthSheet.cell("H1").value("П.Макс");
		fifthSheet.cell("I1").value("П.Мин");

		auto calc_range5 = fifthSheet.range("A1:I1");
		calc_range5.style(Calc_style);
		calc_range5.alignment(center);

		//Шестая книга ///////////////////////////

		sixSheet.cell("A1").value("Код инструмента");
		sixSheet.cell("B1").value("Т.Цена");
		sixSheet.cell("C1").value("Цена Открытия");
		sixSheet.cell("D1").value("Макс");
		sixSheet.cell("E1").value("Мин");
		sixSheet.cell("F1").value("коэф Х");
		sixSheet.cell("G1").value("Х");
		sixSheet.cell("H1").value("П.Макс");
		sixSheet.cell("I1").value("П.Мин");

		auto calc_range6 = sixSheet.range("A1:I1");
		calc_range6.style(Calc_style);
		calc_range6.alignment(center);

		auto nameCell6 = sixSheet.cell("A2");
		makeNoFormulaCellStyle(nameCell6);

		auto priceCell6 = sixSheet.cell("B2");
		makeNoFormulaCellStyle(priceCell6);

		auto openPriceCell6 = sixSheet.cell("C2");
		makeNoFormulaCellStyle(openPriceCell6);

		auto maxCell6 = sixSheet.cell("D2");
		makeNoFormulaCellStyle(maxCell6);

		auto minCell6 = sixSheet.cell("E2");
		makeNoFormulaCellStyle(minCell6);

		auto coefXCell6 = sixSheet.cell("F2");
		coefXCell6.formula("=IF(B2 < C2,1,IF(B2 > C2,2,3))");
		makeCellStyle(coefXCell6);

		auto xCell6 = sixSheet.cell("G2");
		xCell6.formula("=IF(F2=1,(D2+E2+B2+E2)/2,IF(F2 =2,(D2+E2+B2+D2)/2,(D2+E2+B2+B2)/2))");
		makeCellStyle(xCell6);

		auto forcastingMaxCell6 = sixSheet.cell("H2");
		forcastingMaxCell6.formula("=G2 - E2");
		makeCellStyle(forcastingMaxCell6);

		auto forcastingMinCell6 = sixSheet.cell("I2");
		forcastingMinCell6.formula("=G2 - D2");
		makeCellStyle(forcastingMinCell6);
		
		///////////////////////////////////////////////////////////////////
		int rowNum = 5;
		for (const auto& share : stockVec)
		{
			if (share.getSharePrice() == 0) continue;

			std::string name = share.getCompanyName();
			std::string cellName = "A" + std::to_string(rowNum);
			//Первая книга название компании
			auto nameCell1 = newSheet.cell(cellName);
			nameCell1.value(name);
			makeNoFormulaCellStyle(nameCell1);

			//Пятая книга название компании
			std::string cellName5 = "A" + std::to_string(rowNum - 3);
			auto nameCell5 = fifthSheet.cell(cellName5);
			nameCell5.value(name);
			makeNoFormulaCellStyle(nameCell5);
			
			//Вторая книга название компании
			auto nameCell2 = secSheet.cell(cellName);
			nameCell2.value(name);
			makeNoFormulaCellStyle(nameCell2);

			//Первая книга кол-во бумаг в лоте
			int lot = share.getInOneLotShares();
			std::string cellLot = "B" + std::to_string(rowNum);
			auto lotCell1 = newSheet.cell(cellLot);
			lotCell1.value(lot);
			makeNoFormulaCellStyle(lotCell1);
			//Вторая книга кол-во бумаг в лоте
			auto lotCell2 = secSheet.cell(cellLot);
			lotCell2.value(lot);
			makeNoFormulaCellStyle(lotCell2);

			//Первая книга текущая цена
			double price = share.getSharePrice();
			std::string cellPrice = "C" + std::to_string(rowNum);
			auto priceCell1 = newSheet.cell(cellPrice);
			priceCell1.value(price);
			makeNoFormulaCellStyle(priceCell1);

			//Пятая книга текущая цена, цена открытия, макс, мин, коэф Х, Х, Прог.Мах, Прог.Мин
			std::string cellPrice5 = "B" + std::to_string(rowNum - 3);
			auto priceCell5 = fifthSheet.cell(cellPrice5);
			priceCell5.value(price);
			makeNoFormulaCellStyle(priceCell5);

			double openPrice = share.getOpenPrice();
			std::string cellOpenPrice5 = "C" + std::to_string(rowNum - 3);
			auto openPriceCell5 = fifthSheet.cell(cellOpenPrice5);
			openPriceCell5.value(openPrice);
			makeNoFormulaCellStyle(openPriceCell5);

			double maxPrice = share.getMaxPrice();
			std::string cellMaxPrice5 = "D" + std::to_string(rowNum - 3);
			auto maxPriceCell5 = fifthSheet.cell(cellMaxPrice5);
			maxPriceCell5.value(maxPrice);
			makeNoFormulaCellStyle(maxPriceCell5);

			double minPrice = share.getMinPrice();
			std::string cellMinPrice5 = "E" + std::to_string(rowNum - 3);
			auto minPriceCell5 = fifthSheet.cell(cellMinPrice5);
			minPriceCell5.value(minPrice);
			makeNoFormulaCellStyle(minPriceCell5);

			std::string cellXcoef = "F" + std::to_string(rowNum - 3);
			auto coefXCell5 = fifthSheet.cell(cellXcoef);
			coefXCell5.formula("=IF(" + cellPrice5 + " < " + cellOpenPrice5 + ",1,IF(" + cellPrice5 + " > " + cellOpenPrice5 + " ,2,3)) ");
			makeCellStyle(coefXCell5);

			std::string cellX = "G" + std::to_string(rowNum - 3);
			auto xCell5 = fifthSheet.cell(cellX);
			xCell5.formula("=IF(" + cellXcoef + " =1,(" + cellMaxPrice5 + "+" + cellMinPrice5 + " + " + cellPrice5 + "+" + cellMinPrice5 + " )/2,"
				"IF(" + cellXcoef + " =2,(" + cellMaxPrice5 + "+" + cellMinPrice5 + " + " + cellPrice5 + "+" + cellMaxPrice5 + " )/2,"
				"(" + cellMaxPrice5 + "+" + cellMinPrice5 + " + " + cellPrice5 + "+" + cellPrice5 + ")/2))");
			makeCellStyle(xCell5);

			std::string cellForcastingMax = "H" + std::to_string(rowNum - 3);
			auto forcastingMaxCell5 = fifthSheet.cell(cellForcastingMax);
			forcastingMaxCell5.formula("=" + cellX + " - " + cellMinPrice5 + " ");
			makeCellStyle(forcastingMaxCell5);

			std::string cellForcastingMin = "I" + std::to_string(rowNum - 3);
			auto forcastingMinCell5 = fifthSheet.cell(cellForcastingMin);
			forcastingMinCell5.formula("=" + cellX + " - " + cellMaxPrice5 + " ");
			makeCellStyle(forcastingMinCell5);


			//Вторая книга текущая цена
			auto priceCell2 = secSheet.cell(cellPrice);
			priceCell2.value(price);
			makeNoFormulaCellStyle(priceCell2);

			//Первая книга шаг цены
			double priceMove = share.getPriceMove();
			std::string cellPriceMove = "D" + std::to_string(rowNum);
			auto priceMoveCell1 = newSheet.cell(cellPriceMove);
			priceMoveCell1.value(priceMove);
			makeNoFormulaCellStyle(priceMoveCell1);
			//Вторая книга шаг цены
			auto priceMoveCell2 = secSheet.cell(cellPriceMove);
			priceMoveCell2.value(priceMove);
			makeNoFormulaCellStyle(priceMoveCell2);

			//Первая книга стоп цена
			double stopPrice = share.getSupplyPrice();
			if (price == stopPrice) {
				stopPrice = stopPrice - priceMove * 100;
			}
			std::string cellStopPrice = "E" + std::to_string(rowNum);
			auto stopPriceCell1 = newSheet.cell(cellStopPrice);
			stopPriceCell1.value(stopPrice);
			makeNoFormulaCellStyle(stopPriceCell1);
			//Вторая книга стоп цена
			auto stopPriceCell2 = secSheet.cell(cellStopPrice);
			stopPriceCell2.formula("=" + cellPrice + " - " + cellPrice + " * $D$2");
			makeCellStyle(stopPriceCell2);

			//Первая книга профит цена
			double profitPrice = share.getDemandPrice();
			std::string cellProfitPrice = "F" + std::to_string(rowNum);
			xlnt::cell profitPriceCell1 = newSheet.cell(cellProfitPrice);
			profitPriceCell1.value(profitPrice);
			makeNoFormulaCellStyle(profitPriceCell1);
			//Вторая книга профит цена
			auto profitPriceCell2 = secSheet.cell(cellProfitPrice);
			profitPriceCell2.formula("=" + cellPrice + " + " + cellPrice + " * $E$2");
			makeCellStyle(profitPriceCell2);

			//Первая книга цена одного лота
			std::string cellOneLotPrice = "G" + std::to_string(rowNum);
			auto oneLotPriceCell1 = newSheet.cell(cellOneLotPrice);
			oneLotPriceCell1.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			oneLotPriceCell1.formula("=" + cellPrice + " * " + cellLot + "");
			makeCellStyle(oneLotPriceCell1);
			
			//Вторая книга цена одного лота
			auto oneLotPriceCell2 = secSheet.cell(cellOneLotPrice);
			oneLotPriceCell2.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			oneLotPriceCell2.formula("=" + cellPrice + " * " + cellLot + "");
			makeCellStyle(oneLotPriceCell2);

			//Первая книга максимальное кол-во лотов на покупку
			std::string cellMaxLotToBuyFromInvSum = "H" + std::to_string(rowNum);
			auto maxLotToBuyCell1 = newSheet.cell(cellMaxLotToBuyFromInvSum);
			maxLotToBuyCell1.formula("=IF($A$2 < " + cellPrice + " * " + cellLot + " * "
				"ROUNDDOWN(($C$2/((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + "))"
				"/(" + cellLot + " * " + cellPriceMove + "),0), "
			"ROUNDDOWN($A$2/(" + cellPrice + " * " + cellLot + "),0), "
				"ROUNDDOWN(($C$2/((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + "))"
				"/(" + cellLot + " * " + cellPriceMove + "),0))");
			makeCellStyle(maxLotToBuyCell1);
			
			//Вторая книга максимальное кол-во лотов на покупку
			auto maxLotToBuyCell2 = secSheet.cell(cellMaxLotToBuyFromInvSum);
			maxLotToBuyCell2.formula("=IF($A$2 < " + cellPrice + " * " + cellLot + " * "
				"ROUNDDOWN(($C$2/((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + "))"
				"/(" + cellLot + " * " + cellPriceMove + "),0), "
				"ROUNDDOWN($A$2/(" + cellPrice + " * " + cellLot + "),0), "
				"ROUNDDOWN(($C$2/((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + "))"
				"/(" + cellLot + " * " + cellPriceMove + "),0))");
			makeCellStyle(maxLotToBuyCell2);

			//Первая книга максимальный денежный риск в сделке если выйти по стопу
			std::string cellMaxMoneyRist = "I" + std::to_string(rowNum);
			auto maxMoneyRistCell1 = newSheet.cell(cellMaxMoneyRist);
			maxMoneyRistCell1.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxMoneyRistCell1.formula("=" + cellMaxLotToBuyFromInvSum + " * " + cellLot + " * " 
				"" + cellPriceMove + " * ((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + ") ");
			makeCellStyle(maxMoneyRistCell1);
			
			//Вторая книга максимальный денежный риск в сделке если выйти по стопу
			auto maxMoneyRistCell2 = secSheet.cell(cellMaxMoneyRist);
			maxMoneyRistCell2.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxMoneyRistCell2.formula("=" + cellMaxLotToBuyFromInvSum + " * " + cellLot + " * "
				"" + cellPriceMove + " * ((" + cellPrice + " - " + cellStopPrice + ")/" + cellPriceMove + ") ");
			makeCellStyle(maxMoneyRistCell2);

			//Первая книга максимальное использование денежных средств от инв. суммы
			std::string cellMaxUseOfMoney = "J" + std::to_string(rowNum);
			auto maxUseOfMoneyCell1 = newSheet.cell(cellMaxUseOfMoney);
			maxUseOfMoneyCell1.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxUseOfMoneyCell1.formula("=" + cellPrice + " * " + cellLot + " * " + cellMaxLotToBuyFromInvSum + "");
			makeCellStyle(maxUseOfMoneyCell1);
			
			//Вторая книга максимальное использование денежных средств от инв. суммы
			auto maxUseOfMoneyCell2 = secSheet.cell(cellMaxUseOfMoney);
			maxUseOfMoneyCell2.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxUseOfMoneyCell2.formula("=" + cellPrice + " * " + cellLot + " * " + cellMaxLotToBuyFromInvSum + "");
			makeCellStyle(maxUseOfMoneyCell2);

			//Первая книга максимальный остаток денежных средств от инв. суммы
			std::string cellRestMoney = "K" + std::to_string(rowNum);
			auto restMoneyCell1 = newSheet.cell(cellRestMoney);
			restMoneyCell1.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			restMoneyCell1.formula("=$A$2 - " + cellMaxUseOfMoney + "");
			makeCellStyle(restMoneyCell1);
			
			//Вторая книга максимальный остаток денежных средств от инв. суммы
			auto restMoneyCell2 = secSheet.cell(cellRestMoney);
			restMoneyCell2.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			restMoneyCell2.formula("=$A$2 - " + cellMaxUseOfMoney + "");
			makeCellStyle(restMoneyCell2);

			//Первая книга максимальный профит
			std::string cellMaxExpectedProfit = "L" + std::to_string(rowNum);
			auto maxExpectedProfitCell1 = newSheet.cell(cellMaxExpectedProfit);
			maxExpectedProfitCell1.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxExpectedProfitCell1.formula("=" + cellMaxLotToBuyFromInvSum + " * " + cellPriceMove + " * " + cellLot + " * "
				"((" + cellProfitPrice + " - " + cellPrice + ")/" + cellPriceMove + ")");
			makeCellStyle(maxExpectedProfitCell1);
			
			//Вторая книга максимальный профит
			auto maxExpectedProfitCell2 = secSheet.cell(cellMaxExpectedProfit);
			maxExpectedProfitCell2.number_format(xlnt::number_format::number_format("#,##0.00₽"));
			maxExpectedProfitCell2.formula("=" + cellMaxLotToBuyFromInvSum + " * " + cellPriceMove + " * " + cellLot + " * "
				"((" + cellProfitPrice + " - " + cellPrice + ")/" + cellPriceMove + ")");
			makeCellStyle(maxExpectedProfitCell2);
			
			//Первая книга коэффициент риск / профит
			std::string cellRiskProfitRatio = "M" + std::to_string(rowNum);
			auto riskProfitRatioCell1 = newSheet.cell(cellRiskProfitRatio);
			riskProfitRatioCell1.number_format(xlnt::number_format::number_format("#,##0.000"));
			riskProfitRatioCell1.formula("=" + cellMaxMoneyRist + " / " + cellMaxExpectedProfit + "");
			makeCellStyle(riskProfitRatioCell1);
			
			//Вторая книга коэффициент риск / профит
			auto riskProfitRatioCell2 = secSheet.cell(cellRiskProfitRatio);
			riskProfitRatioCell2.number_format(xlnt::number_format::number_format("#,##0.000"));
			riskProfitRatioCell2.formula("=" + cellMaxMoneyRist + " / " + cellMaxExpectedProfit + "");
			makeCellStyle(riskProfitRatioCell2);

			rowNum++;
		}

		createdDoc.save(tableNameCreated);
		return true;
	}

	return false;
}

void fileXLSX::makeNoFormulaCellStyle(xlnt::cell& c)
{

	cell_border_property.style(xlnt::border_style::thin);
	cell_border.side(xlnt::border_side::start, cell_border_property); // left
	cell_border.side(xlnt::border_side::end, cell_border_property); // right
	cell_border.side(xlnt::border_side::top, cell_border_property); // top
	cell_border.side(xlnt::border_side::bottom, cell_border_property); // bottom

	cell_alignment_center.vertical(xlnt::vertical_alignment::center);
	cell_alignment_center.horizontal(xlnt::horizontal_alignment::center);

	//Стиль ячеек под названием CalcNames
	auto calc_font = xlnt::font();
	calc_font.bold(true);
	calc_font.size(12);

	c.fill(xlnt::fill::solid(yellow_color));
	c.border(cell_border);
	c.font(calc_font.color(blue_color));
	c.alignment(cell_alignment_center);
}

void fileXLSX::makeCellStyle(xlnt::cell& c)
{
	cell_border_property.style(xlnt::border_style::thin);
	cell_border.side(xlnt::border_side::start, cell_border_property); // left
	cell_border.side(xlnt::border_side::end, cell_border_property); // right
	cell_border.side(xlnt::border_side::top, cell_border_property); // top
	cell_border.side(xlnt::border_side::bottom, cell_border_property); // bottom

	cell_alignment_center.vertical(xlnt::vertical_alignment::center);
	cell_alignment_center.horizontal(xlnt::horizontal_alignment::center);

	c.border(cell_border);
	c.alignment(cell_alignment_center);
}

void fileXLSX::makeDealBook(std::vector<std::string>& v, double amount, double riskProc) //Создает книгу Deal в таблице
{
	auto dealSheet = createdDoc.create_sheet();
	dealSheet.title("Deals1");//Расчет одного ассета, где можно ставить свои стоп и профит

	//Стиль шрифта 
	auto bold_font = xlnt::font();
	bold_font.bold(true);
	bold_font.size(16);
	//Стиль границы
	xlnt::border border;
	xlnt::border::border_property border_property;
	border_property.style(xlnt::border_style::thin);
	border.side(xlnt::border_side::start, border_property); // left
	border.side(xlnt::border_side::end, border_property); // right
	border.side(xlnt::border_side::top, border_property); // top
	border.side(xlnt::border_side::bottom, border_property); // bottom
	//Выравнивание
	xlnt::alignment center;
	center.vertical(xlnt::vertical_alignment::center);
	center.horizontal(xlnt::horizontal_alignment::center);
	center.wrap(true);
	//Стиль ячеек под названием DealNames
	auto Deal_style = createdDoc.create_style("DealNames");
	Deal_style.font(bold_font);
	Deal_style.alignment(center);
	Deal_style.fill(xlnt::fill::solid(red_color));
	Deal_style.border(border);

	//Стиль ячеек под названием CalcNames
	auto calc_font = xlnt::font();
	calc_font.bold(true);
	calc_font.color(white_color);
	calc_font.size(12);

	auto Calc_style = createdDoc.create_style("CalcNames");
	Calc_style.font(calc_font);
	Calc_style.fill(xlnt::fill::solid(perpule_color));
	Calc_style.border(border);
	Calc_style.alignment(center);

	dealSheet.cell("A1").value("Сумма для инвестиции");
	dealSheet.cell("B1").value("Процент риска");
	dealSheet.cell("C1").value("Сумма риска");

	auto deal_range = dealSheet.range("A1:C1");
	deal_range.style(Deal_style);

	auto cellАmount = dealSheet.cell("A2");
	cellАmount.number_format(xlnt::number_format::number_format("#,##₽"));
	cellАmount.value(amount);
	makeNoFormulaCellStyle(cellАmount);

	auto cellRiskProc = dealSheet.cell("B2");
	cellRiskProc.number_format(xlnt::number_format::percentage_00());
	cellRiskProc.value(riskProc / 100);
	makeNoFormulaCellStyle(cellRiskProc);

	auto cellRiskSum = dealSheet.cell("C2");
	cellRiskSum.number_format(xlnt::number_format::number_format("#,##₽"));
	cellRiskSum.formula("=A2 * B2");
	makeCellStyle(cellRiskSum);

	dealSheet.cell("A4").value("Код инструмента");
	dealSheet.cell("B4").value("Лот");
	dealSheet.cell("C4").value("Т.Цена");
	dealSheet.cell("D4").value("Шаг Цены");
	dealSheet.cell("E4").value("Стоп Цена");
	dealSheet.cell("F4").value("Профит Цена");
	dealSheet.cell("G4").value("Цена лота");
	dealSheet.cell("H4").value("MaxLotFromInvSum");
	dealSheet.cell("I4").value("Макс Денежный Риск");
	dealSheet.cell("J4").value("Макс Исп. Денг.");
	dealSheet.cell("K4").value("Остаток Денег");
	dealSheet.cell("L4").value("Макс Профит");
	dealSheet.cell("M4").value("Коэф. Риск/Профит");

	auto calc_range = dealSheet.range("A4:M4");
	calc_range.style(Calc_style);
	calc_range.alignment(center);

	int numRow = 5;

	for (const std::string comName : v) 
	{
		std::string cellNameStr = "A" + std::to_string(numRow);
		std::string cellLotStr = "B" + std::to_string(numRow);
		std::string cellPriceStr = "C" + std::to_string(numRow);
		std::string cellPriceMoveStr = "D" + std::to_string(numRow);
		std::string cellStopPriceStr = "E" + std::to_string(numRow);
		std::string cellProfitPriceStr = "F" + std::to_string(numRow);

		if (!findComNameInShareVec(comName))
		{
			auto nameCell = dealSheet.cell(cellNameStr);
			makeNoFormulaCellStyle(nameCell);

			auto lotCell = dealSheet.cell(cellLotStr);
			makeNoFormulaCellStyle(lotCell);

			auto priceCell = dealSheet.cell(cellPriceStr);
			makeNoFormulaCellStyle(priceCell);

			auto priceMoveCell = dealSheet.cell(cellPriceMoveStr);
			makeNoFormulaCellStyle(priceMoveCell);

			auto stopPriceCell = dealSheet.cell(cellStopPriceStr);
			makeNoFormulaCellStyle(stopPriceCell);

			auto profitPriceCell = dealSheet.cell(cellProfitPriceStr);
			makeNoFormulaCellStyle(profitPriceCell);

		}
		else {

			for (const StockDataClass& share : stockVec)
			{
				if (comName == share.getCompanyName())
				{

					auto nameCell = dealSheet.cell(cellNameStr);
					nameCell.value(share.getCompanyName());
					makeNoFormulaCellStyle(nameCell);

					auto lotCell = dealSheet.cell(cellLotStr);
					lotCell.value(share.getInOneLotShares());
					makeNoFormulaCellStyle(lotCell);

					auto priceCell = dealSheet.cell(cellPriceStr);
					priceCell.value(share.getSharePrice());
					makeNoFormulaCellStyle(priceCell);

					auto priceMoveCell = dealSheet.cell(cellPriceMoveStr);
					priceMoveCell.value(share.getPriceMove());
					makeNoFormulaCellStyle(priceMoveCell);

					auto stopPriceCell = dealSheet.cell(cellStopPriceStr);
					stopPriceCell.value(share.getSharePrice() * 0.98);
					makeNoFormulaCellStyle(stopPriceCell);

					auto profitPriceCell = dealSheet.cell(cellProfitPriceStr);
					profitPriceCell.value(share.getSharePrice() * 1.05);
					makeNoFormulaCellStyle(profitPriceCell);
					break;
				}

			}
		}


		std::string cellOneLotPriceStr = "G" + std::to_string(numRow);
		auto oneLotPriceCell = dealSheet.cell(cellOneLotPriceStr);
		oneLotPriceCell.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		oneLotPriceCell.formula("= " + cellPriceStr + " * " + cellLotStr + "");
		makeCellStyle(oneLotPriceCell);

		std::string cellMaxLotToBuyStr = "H" + std::to_string(numRow);
		auto maxLotToBuyCell = dealSheet.cell(cellMaxLotToBuyStr);
		maxLotToBuyCell.formula("=IF($A$2 < " + cellPriceStr + " * " + cellLotStr + " * "
			"ROUNDDOWN(($C$2/((" + cellPriceStr + " - " + cellStopPriceStr + ")/ " + cellPriceMoveStr + "))"
			"/(" + cellLotStr + " * " + cellPriceMoveStr + "),0), "
			"ROUNDDOWN($A$2/(" + cellPriceStr + " * " + cellLotStr + "),0), "
			"ROUNDDOWN(($C$2/((" + cellPriceStr + " - " + cellStopPriceStr + ")/" + cellPriceMoveStr + "))"
			"/(" + cellLotStr + " * " + cellPriceMoveStr + "),0))");
		makeCellStyle(maxLotToBuyCell);

		std::string cellMaxMoneyRiskStr = "I" + std::to_string(numRow);
		auto maxMoneyRiskCell = dealSheet.cell(cellMaxMoneyRiskStr);
		maxMoneyRiskCell.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxMoneyRiskCell.formula("=" + cellMaxLotToBuyStr + " * " + cellLotStr + " * "
			"" + cellPriceMoveStr + " * ((" + cellPriceStr + " - " + cellStopPriceStr + ")/" + cellPriceMoveStr + ") ");
		makeCellStyle(maxMoneyRiskCell);

		std::string cellMaxUseOfMoneyStr = "J" + std::to_string(numRow);
		auto maxUseOfMoneyCell = dealSheet.cell(cellMaxUseOfMoneyStr);
		maxUseOfMoneyCell.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxUseOfMoneyCell.formula("=" + cellPriceStr + " * " + cellLotStr + " * " + cellMaxLotToBuyStr + "");
		makeCellStyle(maxUseOfMoneyCell);

		std::string cellRestMoneyStr = "K" + std::to_string(numRow);
		auto restMoneyCell = dealSheet.cell(cellRestMoneyStr);
		restMoneyCell.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		restMoneyCell.formula("=$A$2 - " + cellMaxUseOfMoneyStr + "");
		makeCellStyle(restMoneyCell);

		std::string cellMaxExpectedProfitStr = "L" + std::to_string(numRow);
		auto maxExpectedProfitCell = dealSheet.cell(cellMaxExpectedProfitStr);
		maxExpectedProfitCell.number_format(xlnt::number_format::number_format("#,##0.00₽"));
		maxExpectedProfitCell.formula("=" + cellMaxLotToBuyStr + " * " + cellPriceMoveStr + " * " + cellLotStr + " * "
			"((" + cellProfitPriceStr + " - " + cellPriceStr + ")/" + cellPriceMoveStr + ")");
		makeCellStyle(maxExpectedProfitCell);

		std::string cellRiskProfitRatioStr = "M" + std::to_string(numRow);
		auto riskProfitRatioCell = dealSheet.cell(cellRiskProfitRatioStr);
		riskProfitRatioCell.number_format(xlnt::number_format::number_format("#,##0.000"));
		riskProfitRatioCell.formula("=" + cellMaxMoneyRiskStr + " / " + cellMaxExpectedProfitStr + "");
		makeCellStyle(riskProfitRatioCell);

		numRow++;
	}

	


	createdDoc.save(tableNameCreated);
	
	
	
}

bool fileXLSX::findComNameInShareVec(const std::string com_name)
{
	for (const StockDataClass& share : stockVec) {
		if (com_name == share.getCompanyName()) {
			return true;
		}
	}
	return false;
}

void fileXLSX::calculateOptLotSize(const std::string& comName, double amount, double riskProc)
{
	for (const auto& share : stockVec)
	{
		if (share.getCompanyName() == comName)
		{
			double price = share.getSharePrice();
			int lot = share.getInOneLotShares();
			double lotPrice = share.getLotPrice();
			double priceMove = share.getPriceMove();
			double supply = share.getSupplyPrice();
			double demand = share.getDemandPrice();
			double lotMove = share.getLotMove();

			double movesToStop = (price - supply) / priceMove;
			double priceMoveToStop = (amount * (riskProc/100)) / movesToStop;

			int maxLotInDeal = 0;
			
			if (amount < lotPrice * std::floor(priceMoveToStop / lotMove))
			{
				maxLotInDeal = std::floor(amount / lotPrice);
			}
			else {
				maxLotInDeal = std::floor(priceMoveToStop / lotMove);
			}


			double maxRiskInDeal = maxLotInDeal * lotMove * movesToStop;
			double maxUseOfMoneyInDeal = lotPrice * maxLotInDeal;
			double freeMoneyFromDeal = amount - maxUseOfMoneyInDeal;
			double profitMoves = (demand - price) / priceMove;
			double maxProfitInDeal = profitMoves * maxLotInDeal * lotMove;
			double riskProfitRatio = movesToStop / profitMoves;
			
			printDealData(StockDataOptDialStopProfitClass(comName, amount, riskProc, price, maxLotInDeal, maxRiskInDeal, maxUseOfMoneyInDeal,
			freeMoneyFromDeal, maxProfitInDeal, riskProfitRatio));

		}
	}
}

void fileXLSX::calculateOptLotSizeForAllTable(double amount, double riskProc)
{
	for (const auto& share : stockVec)
	{
		std::string comName = share.getCompanyName();
		double price = share.getSharePrice();
		int lot = share.getInOneLotShares();
		double lotPrice = share.getLotPrice();
		double priceMove = share.getPriceMove();
		double supply = share.getSupplyPrice();
		double demand = share.getDemandPrice();
		double lotMove = share.getLotMove();

		double movesToStop = (price - supply) / priceMove;
		double priceMoveToStop = (amount * (riskProc / 100)) / movesToStop;

		int maxLotInDeal = 0;

		if (amount < lotPrice * std::floor(priceMoveToStop / lotMove))
		{
			maxLotInDeal = std::floor(amount / lotPrice);
		}
		else {
			maxLotInDeal = std::floor(priceMoveToStop / lotMove);
		}


		double maxRiskInDeal = maxLotInDeal * lotMove * movesToStop;
		double maxUseOfMoneyInDeal = lotPrice * maxLotInDeal;
		double freeMoneyFromDeal = amount - maxUseOfMoneyInDeal;
		double profitMoves = (demand - price) / priceMove;
		double maxProfitInDeal = profitMoves * maxLotInDeal * lotMove;
		double riskProfitRatio = movesToStop / profitMoves;

		stockDataOptDialProfitVec.push_back(StockDataOptDialStopProfitClass(comName, amount, riskProc, price, maxLotInDeal, maxRiskInDeal, maxUseOfMoneyInDeal,
			freeMoneyFromDeal, maxProfitInDeal, riskProfitRatio));
	}
	if (createCalculatedTableFromTableData(amount, riskProc)) {

		std::cout << "The calculated data successfully created!" << std::endl;
	}
}


void fileXLSX::printShare(StockCalculateDataClass& data)
{
	data.printOneStockCalData();
}

void fileXLSX::calData(double number, double riskProc)
{
	for (const auto& share : stockVec)
	{
		std::string compName = share.getCompanyName();
		double lotPrice = share.getLotPrice();
		double lotMove = share.getLotMove();
		int lotShare = share.getInOneLotShares();
		double sharePrice = share.getSharePrice();
		double priceMove = share.getPriceMove();

		stockCalculatedVec.push_back(StockCalculateDataClass(compName, number, lotShare, lotPrice, riskProc,
			lotMove, sharePrice, priceMove));
	}
}

int fileXLSX::getColumnNumber(const std::string& columnName)
{	
	auto wks1 = openDoc.workbook().worksheet(sheetName);
	for (int col = 1; col <= wks1.columnCount(); ++col)
	{
		std::string colName = wks1.cell(1, col).value().get<std::string>();
		
		if (columnName.compare(colName) == 0) {
			
			return col;
		}
		
	}
	return 0;
}

int fileXLSX::countFractionalDigits(double number)
{
	double wholePart;
	double fractionalPart = std::modf(number, &wholePart);

	std::stringstream ss;
	ss << fractionalPart;

	std::string str = ss.str();
	int count = 0;
	for (char c : str) {
		if (c != '0' && c != '.') {
			count++;
		}
	}
	return count;
}
