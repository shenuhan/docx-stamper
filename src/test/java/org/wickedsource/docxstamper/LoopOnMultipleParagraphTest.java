package org.wickedsource.docxstamper;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;
import org.wickedsource.docxstamper.context.Character;
import org.wickedsource.docxstamper.context.ListCharacter;
import org.wickedsource.docxstamper.processor.repeat.ILoopProcessor;
import org.wickedsource.docxstamper.processor.repeat.LoopProcessor;

public class LoopOnMultipleParagraphTest extends AbstractDocx4jTest {
    @Test
    public void processorExpressionsInCommentsAreResolved() throws Docx4JException, IOException {
    	ListCharacter peopleContext = new ListCharacter();
    	for (final String name : new String[]{"Homer", "Bart", "Lisa", "Marge"}) {
    		peopleContext.getPeople().add(new Character(name,name));

    	}
        InputStream template = getClass().getResourceAsStream("LoopOnMultipleParagraphTest.docx");
        DocxStamperConfiguration config = new DocxStamperConfiguration().addCommentProcessor(ILoopProcessor.class, new LoopProcessor());
        WordprocessingMLPackage document = stampAndLoad(template, peopleContext, config);
        document.save(File.createTempFile("LoopOnMultipleParagraphTest", ".docx"));
    }
}
