package test;

import java.io.IOException;
import java.util.Scanner;

import test.readFromExcel;

public class main {
	public static void main(String[] args) throws IOException {


		String inputName = null;
		String outputName = null;
	
		System.out.println("drop to excel file you want to read\n");
		Scanner scanIn = new Scanner(System.in);
		inputName = scanIn.nextLine();
		inputName = inputName.replaceAll("\\s", "");
		System.out.println("enter to destination path  \n");
		Scanner scanOu = new Scanner(System.in);
		outputName = scanOu.nextLine();
		outputName = outputName.replaceAll("\\s", "");
		readFromExcel test = new readFromExcel(inputName, outputName);

		test.read();

	}
}
