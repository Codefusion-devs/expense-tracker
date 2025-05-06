package com.expense.service;

import java.io.IOException;
import java.io.InputStream;

import com.expense.model.Expense;

public interface ExpenseServiceI {
    void addExpense(Expense expense, String username) throws IOException;
    InputStream getExpenseFile(String username) throws IOException;
    

}
