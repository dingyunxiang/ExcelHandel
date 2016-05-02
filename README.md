# ExcelHandel
用于处理Excel与ArrayList之间的转化

使用POI解析Excel

        ExcelHandle<testClass> handle = new ExcelHandle<testClass>(testClass.class);


        List<testClass> list1 = new ArrayList<testClass>();

        for(int i=0;i<10;i++){
            list1.add(new testClass(i,i));
        }

        String[][] a = {{"name","姓名"},{"age","年龄"}};

        Map<String,String> map = new HashMap<String,String>();
        for(int i=0;i<a.length;i++){
            map.put(a[i][1],a[i][0]);
        }
        
       // System.out.println(map.get("姓名"));

        File f = new File("workbook.xls");
        System.out.println("file:"+f);
        List<testClass> list = handle.excelToList(f,a,true);

        for(testClass t:list){
            System.out.println(t.getName());
        }
