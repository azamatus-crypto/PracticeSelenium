import java.io.IOException;
import java.util.ArrayList;

public class TestSample {
    public static void main(String[] args) throws IOException {
       Examples examples=new Examples();
       ArrayList arrayList=examples.getInfoFromExel("Product Name");
        System.out.println(arrayList.get(0));
        System.out.println(arrayList.get(1));
        System.out.println(arrayList.get(2));
        Practice practice=new Practice();






    }
}
