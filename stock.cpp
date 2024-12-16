#include"stock.h"
//#include <boost/nowide/iostream.hpp>

void StockDataClass::printStockData() const
{
	std::cout << stockInfo.companyCode << ": share_price-> "  << stockInfo.sharePrice <<
		" lot -> " << stockInfo.lot <<  " lot_price-> " << stockInfo.lotPrice << 
		" price_move -> "<< stockInfo.priceMove << 
		" lot_move -> " << stockInfo.lotMove << 
		" supply_price -> " << stockInfo.supplyPrice <<
		" demand_price -> " << stockInfo.demandPrice <<
		std::endl;
}

void StockCalculateDataClass::printStockCalculatedData() const
{
	std::cout << dialStockData.companyCode << ": DS-> " << dialStockData.accNum <<
		" LOTS -> " << dialStockData.lotToBuy << " Risk%-> " << dialStockData.procToRist * 100 <<"%" <<
		" RiskDS -> " << dialStockData.riskAmount <<
		" MovesToStop -> " << dialStockData.priceMoveToStop << 
		" StopPrice -> " << dialStockData.stopPrice <<
		std::endl;
}

void StockCalculateDataClass::printOneStockCalData() const
{
	std::cout << "Company code: " << dialStockData.companyCode << std::endl;
	std::cout << "Investing pile: " << dialStockData.accNum << std::endl;
	std::cout << "The amount of lots to buy: " << dialStockData.lotToBuy <<
		" total number  "<< dialStockData.oneLot * dialStockData.lotToBuy << " shares " << std::endl;
	std::cout << "Risk in procents: " << dialStockData.procToRist * 100 << "%" << std::endl;
	std::cout << "Money to risk: " << dialStockData.riskAmount << std::endl;
	std::cout << "Price Moves until reach the stop: " << dialStockData.priceMoveToStop << std::endl;
	std::cout << "Current Price: " << dialStockData.currentPrice << std::endl;
	std::cout << "Stop Price: " << dialStockData.stopPrice << std::endl;
	std::cout << std::endl;
}

void StockDataOptDialStopProfitClass::printStockDataOptDialStopProfit() 
{
	std::cout << "Company Name: " << stockData.companyCode << std::endl;
	std::cout << "Invest Amount: " << std::fixed << std::setprecision(2) << stockData.investAmount << std::endl;
	std::cout << "Risk Procent: " << stockData.riskProcent << "%" << std::endl;
	std::cout << "Share price: " << std::fixed << std::setprecision(countFractionalDigits(stockData.companyPrice)) << stockData.companyPrice << std::endl;
	std::cout << "Max Lots can buy: " << std::fixed << std::setprecision(2) << stockData.maxLotInDeal << std::endl;
	std::cout << "Max Loss: " << stockData.maxRiskInDeal << std::endl;
	std::cout << "Max Use of invest amount: " << stockData.maxUseOfMoneyInDeal << std::endl;
	std::cout << "Money not used in possition: " << stockData.freeMoneyFromDeal << std::endl;
	std::cout << "Max Profit: " << stockData.maxProfitInDeal << std::endl;
	std::cout << "Risk / Profit Ratio: " << stockData.riskProfitRatio << std::endl;
}

int StockDataOptDialStopProfitClass::countFractionalDigits(double number)
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


