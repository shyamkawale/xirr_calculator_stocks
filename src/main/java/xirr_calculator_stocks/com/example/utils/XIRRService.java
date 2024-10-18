package xirr_calculator_stocks.com.example.utils;

import java.util.*;

import org.apache.commons.math3.analysis.UnivariateFunction;
import org.apache.commons.math3.analysis.solvers.BrentSolver;

public class XIRRService {
    // XIRR Calculation Method (using Brent's Method)
    public static double calculateXIRR(List<Map<String, Object>> cashFlows, double guess) {
        BrentSolver solver = new BrentSolver(1e-6, 1e-6);  // precision, tolerance

        // Define the XIRR function using UnivariateFunction
        UnivariateFunction xirrFunction = (double rate) -> xirrFunction(cashFlows, rate);

        // Solve for the XIRR using the BrentSolver
        return solver.solve(1000, xirrFunction, -0.9999, 10, guess); // range for rates (-99.99% to very high)
    }

    // Function to calculate the XIRR formula for a given rate (x)
    private static double xirrFunction(List<Map<String, Object>> cashFlows, double rate) {
        double result = 0.0;
        Date startDate = (Date) cashFlows.get(0).get("date");

        for (Map<String, Object> flow : cashFlows) {
            Date date = (Date) flow.get("date");
            double cashFlow = (double) flow.get("cashFlow");

            // Difference in days between the cash flow date and the start date
            double days = (date.getTime() - startDate.getTime()) / (1000.0 * 60 * 60 * 24);
            result += cashFlow / Math.pow(1.0 + rate, days / 365.0);
        }

        return result;
    }
}
