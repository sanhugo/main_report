package ru.doc;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelManager {
    public void writeProtocol(int num, File file, XSSFWorkbook book, XSSFRow newRow) throws IOException, InvalidFormatException {
        String name = nameOfSubject(file.getName());// у 4 класса только две предметные олимпиады - вводится условие обработки файлов
        System.out.println(name);
        if (num>4 || name.equals("Математика") || name.equals("Русский язык")){
            int i=3;
            while (newRow.getCell(i)!=null && newRow.getCell(i).getCellType() != CellType.BLANK) //проверка на существование ячейки и на пустоту
                i++;
            newRow.createCell(i); //создание ячейки
            newRow.getCell(i).setCellValue(name);
            //прописываем дисциплину, по которой вносим данные, однако у 4 класса
            String listName = String.format("%d класс",num);
            copyNames(listName,book,file,i); //копируем из предметного протокола в общий
        }
    }

    public void makeClass(int num, XSSFWorkbook book, String src) throws IOException, InvalidFormatException {
        XSSFSheet newSheet =book.createSheet(String.format("%d класс",num));
        XSSFRow newRow = newSheet.createRow(1);
        XSSFCell[] cells = new XSSFCell[3];
        for (int d=0;d<3;d++)
            cells[d] = newRow.createCell(d);
        cells[0].setCellValue("№");
        cells[1].setCellValue("ФИО");
        cells[2].setCellValue("Класс");
        File r = new File (src);
        File[] w = r.listFiles();
        assert w != null;
        for (File file: w){ //запуск обработки предметных протоколов
            String name = file.getName();
            if (FilenameUtils.getExtension(name).equals("xlsm")){
                writeProtocol(num,file,book, newRow);
            }
        }
        newSheet.setColumnWidth(1,6000);
    }

    public void generateTotal(String path, String year, String src) throws IOException, InvalidFormatException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream(path + String.format("участие во ВОШ %s.xlsx",year));
        for (int i=4;i<12;i++)
            makeClass(i,book,src);
        createTotal(book);
        book.write(fileOut);
        fileOut.close();
    }

    public void createTotal(XSSFWorkbook book) { // сращиваем 5-11 классы в один протокол
        book.setSheetName(book.getSheetIndex("4 класс"), "Школьный 4 класс");
        XSSFSheet total = book.createSheet("Школьный 5-11");
        int totalIter=2;
        for (int i=5;i<12;i++)
        {
            XSSFSheet current = book.getSheet(String.format("%d класс",i));
            if (i==5)
            {
                copyRows(current,total,1,1);
            }
            int g=2;
            while (current.getRow(g)!=null)
            {
                copyRows(current, total, g, totalIter);
                total.getRow(totalIter).createCell(0).setCellValue(totalIter-1);
                totalIter++;
                g++;
            }
        }
        createBottom(book,total,totalIter);
        XSSFSheet sheet = book.getSheet("Школьный 4 класс");
        int rowTotal = sheet.getLastRowNum();
        if ((rowTotal > 0) || (sheet.getPhysicalNumberOfRows() > 0)) {
            rowTotal++;
        }
        for (int i=2;i<rowTotal;i++)
        {
            sheet.getRow(i).createCell(0).setCellValue(i-1);
        }
        createBottom(book,sheet,rowTotal);
    }

    public void createBottom(XSSFWorkbook book,XSSFSheet total, int totalIter) {
        int columns = total.getRow(1).getLastCellNum();
        String columnNum=getColumn(columns);
        for (int i=2;i<totalIter;i++)
        {
            XSSFCell cell=total.getRow(i).createCell(columns,CellType.FORMULA);
            String a = String.format("COUNTA(D%d:%s%d)",i+1,columnNum,i+1);
            //System.out.println(a);
            cell.setCellFormula(a);
        }
        XSSFRow totalRow1 = total.createRow(totalIter);
        CellStyle s = book.createCellStyle();
        XSSFFont font =book.createFont();
        font.setBold(true);
        s.setFont(font);
        totalRow1.createCell(2).setCellStyle(s);
        totalRow1.createCell(2).setCellValue("Участники");
        //System.out.println();
        for (int i=3;i<=columns;i++)
        {
            String row;
            if (i<columns)
                row=getColumn(i+1);
            else row=getColumn(i);
            XSSFCell cell2 = totalRow1.createCell(i);
            String a;
            if (i<columns)
            {
                a = String.format("COUNTA(%s3:%s%d)", row, row, totalIter);
            }
            else
            {
                a = String.format("SUM(D%d:%s%d)", totalIter + 1, row, totalIter + 1);
            }
            //System.out.println(a);
            cell2.setCellFormula(a);
            cell2.setCellStyle(s);
        }
        totalRow1=total.createRow(totalIter+1);
        totalRow1.createCell(2).setCellStyle(s);
        totalRow1.createCell(2).setCellValue("Призеры");
        for (int i=3;i<=columns;i++)
        {
            String row;
            if (i<columns)
                row=getColumn(i+1);
            else row=getColumn(i);
            XSSFCell cell2 = totalRow1.createCell(i);
            cell2.setCellStyle(s);
            String a;
            if (i<columns){
                a = String.format("COUNTIF(%s3:%s%d, \"пр\")", row, row, totalIter);
            }
            else
            {
                a = String.format("SUM(D%d:%s%d)", totalIter + 2, row, totalIter + 2);
            }
            //System.out.println(a);
            cell2.setCellFormula(a);
        }
        totalRow1=total.createRow(totalIter+2);
        totalRow1.createCell(2).setCellStyle(s);
        totalRow1.createCell(2).setCellValue("Победители");
        for (int i=3;i<=columns;i++)
        {
            String row;
            if (i<columns)
            row=getColumn(i+1);
            else row=getColumn(i);
            XSSFCell cell2 = totalRow1.createCell(i);
            cell2.setCellStyle(s);
            String a;
            if (i<columns)
                a=String.format("COUNTIF(%s3:%s%d, \"поб\")",row,row,totalIter);
            else a = String.format("SUM(D%d:%s%d)", totalIter + 3, row, totalIter + 3);
            //System.out.println(a);
            cell2.setCellFormula(a);
        }
        total.createRow(0).createCell(0).setCellValue("Участие в школьном этапе Всероссийской олимпиады школьников");
    }

    public String getColumn(int i)
    {
        String row="";
        if (i>=26) row+="A";
        char sym=(char)('A'+(i-1)%26);
        row+=sym;
        return row;
    }
    public void copyNames(String sheetName, XSSFWorkbook book, File file, int columnNumber) throws IOException, InvalidFormatException {
        XSSFWorkbook copiedBook = new XSSFWorkbook(file);
        XSSFSheet copiedBookSheet = copiedBook.getSheet(sheetName);
        //XSSFSheet writingSheet = book.getSheet(sheetName);
        //перебор строк в файле с 8 строки до того, пока не null-строка
        //прибавляем ещё и к номеру новой строки на итерациях
        //три ячейки - в одну (ФИО), в номере колонки ставим +
        //System.out.println(copiedBookSheet.getSheetName());
        int i=2;
        XSSFRow b = copiedBookSheet.getRow(i+13);
        while (b!=null && b.getCell(1)!=null && b.getCell(1).getCellType() !=CellType.BLANK)
        //если строка null, то ячейка не проверится; если ячейка null, то на пустоту она не проверится
        {
            String fio = String.format("%s %s %s",b.getCell(2).getStringCellValue().trim(),b.getCell(3).getStringCellValue().trim(),b.getCell(4).getStringCellValue().trim());
            boolean flag = true;
            int g=1;
            int classnum =(int) (b.getCell(10).getNumericCellValue());
            if (classnum==0) System.out.println(file.getName());
            XSSFSheet writingSheet;
            if (book.getSheet(String.format("%d класс",classnum))!=null) {
                writingSheet= book.getSheet(String.format("%d класс",classnum));
            } else
                writingSheet=book.createSheet(String.format("%d класс", classnum));
            XSSFRow a = writingSheet.getRow(g);
            while (a!=null)
            {
                //System.out.println(a.getCell(1));
                if (a.getCell(1).getStringCellValue().equals(fio))
                {
                    flag=false;
                    //System.out.println(g+" Cycle was finished!");
                    break;
                }
                g++;
                a=writingSheet.getRow(g);
            }
            if(!flag)
            {
                a=writingSheet.getRow(g);
                XSSFCell sub=a.createCell(columnNumber);
                sub.setCellValue("+");
            }
            else
            {
                a = writingSheet.createRow(g);
                XSSFCell[] cells = new XSSFCell[4];
                for (int f=0;f<4;f++)
                    if (f==3)
                        cells[f]=a.createCell(columnNumber);
                    else
                        cells[f] = a.createCell(f);
                cells[1].setCellValue(fio);
                cells[3].setCellValue("+");
                cells[2].setCellValue(classnum);
            }
            i++;
            b = copiedBookSheet.getRow(i+13);
        }
        System.out.println(file.getName()+" "+ sheetName);
        copiedBook.close();
    }

    public String nameOfSubject(String name) //парсим название предмета и исправляем
    {
        String a="";
        String s = name.toLowerCase();
        String[] d = {"астрономия","биология","география","информатика","история","китайский язык","литература","математика","мхк",
                "обж","обществознание","право","физика","химия","русский язык","физическая культура", "технология", "английский язык"};
        for (String value : d)
            if (s.contains(value)) {
                a += value;
                break;
            }
        if (!a.isEmpty()) {
            if (a.equals("обж") || a.equals("мхк")) return a.toUpperCase();
            else return a.substring(0, 1).toUpperCase() + a.substring(1);//название предмета с большой буквы
        }
        else return a;
    }
    public void copyRows(XSSFSheet current, XSSFSheet total,int srcnum, int destnum)
    {
        XSSFRow sourceRow = current.getRow(srcnum);
        XSSFRow newRow = total.createRow(destnum);
        // Loop through source columns to add to new row
        for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
            XSSFCell oldCell = sourceRow.getCell(j);
            XSSFCell newCell = newRow.createCell(j);
            if (oldCell == null) {
                continue;
            }
            newCell.setCellType(oldCell.getCellType());
            // Set the cell data value
            switch (oldCell.getCellType()) {
                case BLANK->// Cell.CELL_TYPE_BLANK:
                        newCell.setCellValue(oldCell.getStringCellValue());
                case BOOLEAN-> newCell.setCellValue(oldCell.getBooleanCellValue());
                case FORMULA-> newCell.setCellFormula(oldCell.getCellFormula());
                case NUMERIC->newCell.setCellValue(oldCell.getNumericCellValue());
                case STRING-> newCell.setCellValue(oldCell.getRichStringCellValue());
                default-> {
                }
            }
        }
    }
}