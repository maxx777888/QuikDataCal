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
	
	std::vector<std::string> assets_vec = {"SBER", "GAZP","LKOH", "VTBR","SNGSP",
		"ROSN", "RTKMP", "MRKP", "FEES", "GMKN","NVTK","IRAO"};

	double dealSum = 1000000.0;
	double riskProc = 5.00;
	fileXLSX f("QuikASdata.xlsx", "Текущие торги", "QuikDataReady.xlsx");
	f.calculateOptLotSizeForAllTable(dealSum, riskProc);
	f.makeDealBook(assets_vec, dealSum, riskProc);


	
	return 0;
}
