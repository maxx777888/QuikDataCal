#pragma once
#include <OpenXLSX.hpp>
#include <xlnt/xlnt.hpp>
#include <iostream>
#include <cmath>
#include <numeric>
#include <iomanip>
#include <sstream>
#include <vector>
#include <string>
#include <algorithm>
#include "stock.h"



class fileXLSX
{
public:

	fileXLSX(const std::string& fileName_, const std::string& sheetName_);
	~fileXLSX();
	std::vector<StockDataClass> getAllDataFromFile();
	void printStartData();//Печатает данные взятые из таблицы
	//Печатает расчитанные данные взависимости от инв. суммы и процента риска. Стоп рассчитывается от суммы риска
	void printCalculatedData(double number, double riskProc); 
	//Печатает расчитанные данные одной бумаги взависимости от инв. суммы и процента риска. Стоп рассчитывается от суммы риска
	void printShareCalData(const std::string& comName, double amount, double riskProc);

	//Расчитывает оптимальный лот для стопа и профита в сделке 
	void calculateOptLotSize(const std::string& comName, double amount, double riskProc);

	//Расчитывает оптимальный лот для стопа и профита из данных в таблице
	void calculateOptLotSizeForAllTable(double amount, double riskProc);

	void makeDealBook(std::vector<std::string>& v, double amount, double riskProc); //Создает книгу Deal в таблице
	
private:
	std::string fileName = "";
	std::string sheetName = "";
	////////////////////// Название столбцов в таблице
	std::string sharePriceNameStr = "Цена послед.";
	std::string shareCodeNameStr = "Код инструмента";
	std::string shareLotNameStr = "Лот";
	std::string priceMoveNameStr = "Шаг цены";
	std::string bestSupplyNameStr = "Лучш. пред";
	std::string bestDemandNameStr = "Лучш. спрос";
	std::string minPriceStr = "Мин. цена";
	std::string maxPriceStr = "Макс. цена";
	std::string assetVolumeStr = "Оборот";
	std::string openPriceStr = "Цена закр.";
	/////////////////////////////////////////////
	void calData(double number, double riskProc);
	void printShare(StockCalculateDataClass& data);
	//////////////////////////////////////////////////
	xlnt::rgb_color yellow_color = xlnt::rgb_color(255, 253, 79);
	xlnt::rgb_color red_color = xlnt::rgb_color(255, 102, 102);
	xlnt::rgb_color perpule_color = xlnt::rgb_color(153, 0, 255);
	xlnt::rgb_color white_color = xlnt::rgb_color(255, 255, 255);
	xlnt::rgb_color blue_color = xlnt::rgb_color(0, 0, 139);
	//////////////////////////////////////////////////////////
	xlnt::border cell_border;
	xlnt::border::border_property cell_border_property;

	xlnt::alignment cell_alignment_center;
	

	OpenXLSX::XLDocument openDoc;
	xlnt::workbook createdDoc;

	std::vector<StockDataClass> stockVec;
	std::vector<StockCalculateDataClass> stockCalculatedVec;
	std::vector<StockDataOptDialStopProfitClass> stockDataOptDialProfitVec;

	int getColumnNumber(const std::string& columnName);//По названию метод находит номер столбца
	int countFractionalDigits(double number); //Подсчитывает кол-во цифр после запятой в дробном числе
	
	//Выводит на экран информацию по сделке
	void printDealData(StockDataOptDialStopProfitClass data);

	bool createCalculatedTableFromTableData(double amount, double riskProc);

	void makeNoFormulaCellStyle(xlnt::cell& c);//Делает стиль ячейки без формулы
	void makeCellStyle(xlnt::cell& c);//Делает стиль ячейки с формулой

	bool findComNameInShareVec(const std::string com_name);//Возвращает правда если находит код компании в векторе активов
};

