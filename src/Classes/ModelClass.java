package Classes;

import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;

public class ModelClass {
    private final SimpleIntegerProperty num;
    private final SimpleStringProperty col1;
    private final SimpleStringProperty col2;
    private final SimpleStringProperty col3;
    private final SimpleStringProperty col4;
    private final SimpleStringProperty col5;
    private final SimpleStringProperty col6;


    public ModelClass(int num, String col1, String col2, String col3, String col4, String col5,String col6){
        this.num = new SimpleIntegerProperty(num);
        this.col1 = new SimpleStringProperty(col1);
        this.col2 = new SimpleStringProperty(col2);
        this.col3 = new SimpleStringProperty(col3);
        this.col4 = new SimpleStringProperty(col4);
        this.col5 = new SimpleStringProperty(col5);
        this.col6 = new SimpleStringProperty(col5);


    }

    public int getNum() {
        return num.get();
    }
    public void setNum(int num) {
        this.num.set(num);
    }

    public String getCol1() {
        return col1.get();
    }
    public void setCol1(String col1) {
        this.col1.set(col1);
    }

    public String getCol2() {
        return col2.get();
    }
    public void setCol2(String col2) {
        this.col2.set(col2);
    }

    public String getCol3() {
        return col3.get();
    }
    public void setCol3(String col3) {
        this.col3.set(col3);
    }

    public String getCol4() {
        return col4.get();
    }
    public void setCol4(String col4) {
        this.col4.set(col4);
    }


    public String getCol5() {
        return col5.get();
    }
    public void setCol5(String col5) {
        this.col5.set(col5);
    }

    public String getCol6() {
        return col6.get();
    }
    public void setCol6(String col6) {
        this.col6.set(col6);
    }

}
