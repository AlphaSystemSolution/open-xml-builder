package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.testng.annotations.Test;

public class TableAdapterTest extends CommonTest {

    @Override
    protected String getFileName() {
        return "table-adapter-test.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return WmlPackageBuilder.createPackage().getPackage();
    }

    @Test
    public void testTableAdapter() {
        var tableAdapter = new TableAdapter().startTable(25.0, 8.0, 17.0, 17.0, 8.0, 25.0);

        tableAdapter.startRow()
                .addColumn(0, 6, WmlAdapter.getParagraph("Column spans all grid spans."))
                .endRow();

        tableAdapter.startRow()
                .addColumn(0, 1, WmlAdapter.getParagraph("1"))
                .addColumn(1, 2, WmlAdapter.getParagraph("2"))
                .addColumn(3, 2, WmlAdapter.getParagraph("3"))
                .addColumn(5, 1, WmlAdapter.getParagraph("4"))
                .endRow();

        tableAdapter.startRow()
                .addColumn(0, 2, WmlAdapter.getParagraph("Column 1 of row with 3 columns"))
                .addColumn(2, 2, WmlAdapter.getParagraph("Column 2 of row with 3 columns"))
                .addColumn(4, 2, WmlAdapter.getParagraph("Column 3 of row with 3 columns"))
                .endRow();

        getMainDocumentPart().addObject(tableAdapter.getTable());
    }
}
