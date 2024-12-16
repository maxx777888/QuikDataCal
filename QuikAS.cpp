#pragma execution_character_set("utf-8")
//#include <OpenXLSX.hpp>
#include <iostream>
#include <cmath>
#include <numeric>
#include <iomanip>
#include <vector>
#include <algorithm>
#include "stock.h"
#include "menu.h"
#include"dataInputXLSX.h"


int main()
{
	//std::locale::global(std::locale("ru_RU.UTF-8"));
	std::vector<std::string> assets_vec = {"SBER", "GAZP","LKOH", "VTBR","SNGSP",
		"ROSN", "RTKMP", "MRKP", "FEES", "GMKN","NVTK"};

	double dealSum = 1000000.0;
	double riskProc = 5.00;
	fileXLSX f("QuikAS1.xlsx", "Все Акции Московской биржи #2");
	f.calculateOptLotSizeForAllTable(dealSum, riskProc);
	f.makeDealBook(assets_vec, dealSum, riskProc);


	//f.printStartData();
	//f.printCalculatedData(1000000.00, 2.00);
	/*f.printShareCalData("GAZP", dealSum, riskProc);
	f.printShareCalData("SNGSP", dealSum, riskProc);
	f.printShareCalData("FEES", dealSum, riskProc);
	f.printShareCalData("MRKP", dealSum, riskProc);*/
	//f.calculateOptLotSize("MSNG", dealSum, riskProc);

	
	return 0;
}
