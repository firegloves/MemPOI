/**
 * created by firegloves
 */

package it.firegloves.mempoi.manager;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.testutil.PrivateAccessHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.greaterThan;
import static org.hamcrest.Matchers.instanceOf;
import static org.junit.Assert.assertNotNull;

public class FileManagerTest {

    private FileManager fileManager;
    private WorkbookConfig wbConfig;

    @Mock
    private Workbook wb;

    @Before
    public void prepare() {

        this.wbConfig = new WorkbookConfig().setWorkbook(new SXSSFWorkbook());
        this.fileManager = new FileManager(this.wbConfig);
    }

    /******************************************************************************************************************
     *                          writeFile
     *****************************************************************************************************************/

    @Test
    public void writeFileTest() {

        File file = new File("temp.xlsx");

        this.fileManager.createFinalFile(file);

        Assert.assertTrue("file generated", file.exists());

        file.delete();
    }

    @Test(expected = MempoiException.class)
    public void writeFileTest_nullWorkbook() {

        new FileManager(new WorkbookConfig()).createFinalFile(new File("/not_existing/temp.xlsx"));
    }

    @Test(expected = MempoiException.class)
    public void writeFileTest_invalidFilePath() throws Throwable {

        this.fileManager.createFinalFile(new File("/not_existing/temp.xlsx"));
    }

    @Test(expected = MempoiException.class)
    public void writeFileTest_invalidFilePathAndNullWorkbook() {

        this.fileManager.createFinalFile(new File("/not_existing/temp.xlsx"));
    }

    @Test(expected = MempoiException.class)
    public void writeFileTest_errorWriting() throws Throwable {

        try {

            MockitoAnnotations.initMocks(this);
            Mockito.doThrow(new RuntimeException()).when(wb).write(Mockito.any());

            FileManager fileManager = new FileManager(new WorkbookConfig().setWorkbook(wb));

            fileManager.createFinalFile(new File("temp.xlsx"));

        } catch (Exception e) {
            throw e.getCause();
        }
    }


    /******************************************************************************************************************
     *                          writeTempFile
     *****************************************************************************************************************/

    @Test
    public void writeTempFile() {


        File file = this.fileManager.writeTempFile();

        Assert.assertNotNull("file not null", file);
        Assert.assertTrue("file exists", file.exists());
    }

    /******************************************************************************************************************
     *                          writeToByteArray
     *****************************************************************************************************************/

    @Test
    public void writeToByteArray() throws Exception {

        Method m = PrivateAccessHelper.getAccessibleMethod(FileManager.class, "writeToByteArray");
        Object o = m.invoke(this.fileManager);

        assertNotNull("not null", o);
        assertThat("instance of byte[]", o, instanceOf(byte[].class));

        byte[] bytes = (byte[]) o;

        assertThat("byte array not empty", bytes.length, greaterThan(0));
    }


    @Test(expected = InvocationTargetException.class)
    public void writeToByteArrayNullWorkbook() throws Exception {

        FileManager fileManager = new FileManager(new WorkbookConfig());

        Method m = PrivateAccessHelper.getAccessibleMethod(FileManager.class, "writeToByteArray");
        m.invoke(fileManager);
    }


    /******************************************************************************************************************
     *                          close workbook
     *****************************************************************************************************************/

    @Test
    public void closeWorkbook() throws Exception {

        Method m = PrivateAccessHelper.getAccessibleMethod(FileManager.class, "closeWorkbook");
        m.invoke(this.fileManager);
    }

    @Test(expected = MempoiException.class)
    public void closeWorkbookNullWorkbook() throws Throwable {

        FileManager fileManager = new FileManager(new WorkbookConfig());

        Method m = PrivateAccessHelper.getAccessibleMethod(FileManager.class, "closeWorkbook");

        try {
            m.invoke(fileManager);
        } catch (Exception e) {
            throw e.getCause();
        }
    }


    @Test(expected = MempoiException.class)
    public void closeWorkbookErrorClosing() throws Throwable {

        MockitoAnnotations.initMocks(this);
        Mockito.doThrow(new RuntimeException()).when(wb).close();

        FileManager fileManager = new FileManager(new WorkbookConfig().setWorkbook(wb));

        Method m = PrivateAccessHelper.getAccessibleMethod(FileManager.class, "closeWorkbook");

        try {
            m.invoke(fileManager);
        } catch (Exception e) {
            throw e.getCause();
        }
    }
}
