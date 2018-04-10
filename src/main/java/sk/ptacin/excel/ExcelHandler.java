package sk.ptacin.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.BeanFactory;
import org.springframework.context.support.GenericApplicationContext;
import sk.ptacin.excel.config.SpringContext;

/**
 * Created by Michal Ptacin (michal.ptacin@icz.sk) on 31. 5. 2016.
 */
public class ExcelHandler {

    private static final Logger log = LoggerFactory.getLogger(ExcelHandler.class);


    public static void main(String[] args) throws Exception {
        log.info("Verzia:{}",ExcelHandler.class.getPackage().getImplementationVersion());
        GenericApplicationContext ctx = SpringContext.getContext();
        BeanFactory factory = ctx;
        CopyProcessor copyProcessorInstance = (CopyProcessor) factory.getBean("copyProcessor");
        copyProcessorInstance.startCopying();

    }

}
