package org.example;

import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.util.List;

public class Main implements PropertyChangeListener {
    public static void main(String[] args) throws Exception {
        Main app = new Main();
        ReadLargeExcelFile readLargeExcelFile = new ReadLargeExcelFile();
        readLargeExcelFile.addPropertyChangeListener(app);
        readLargeExcelFile.processSheets("/temp/largeFile.xlsx");
    }

    @Override
    public void propertyChange(PropertyChangeEvent evt) {
        List<String> cols = (List<String>) evt.getNewValue();
        System.out.println(cols);
    }
}