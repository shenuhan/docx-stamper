package org.wickedsource.docxstamper.proxy;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.wickedsource.docxstamper.api.commentprocessor.ICommentProcessor;
import org.wickedsource.docxstamper.api.coordinates.ParagraphCoordinates;
import org.wickedsource.docxstamper.api.coordinates.RunCoordinates;
import org.wickedsource.docxstamper.api.coordinates.TableCoordinates;
import org.wickedsource.docxstamper.processor.CommentProcessorRegistry;

public class TestImpl implements ITestInterface, ICommentProcessor {

	@Override
	public String returnString(String string) {
		return string;
	}

	@Override
	public void commitChanges(WordprocessingMLPackage document) {

	}

	@Override
	public void setCurrentParagraphCoordinates(ParagraphCoordinates coordinates) {

	}

	@Override
	public void setCurrentRunCoordinates(RunCoordinates coordinates) {

	}

	@Override
	public void reset() {

	}

	@Override
	public void setRegistry(CommentProcessorRegistry registry) {

	}

	@Override
	public void onParagraphe(ParagraphCoordinates paragrapheCoordinates) {

	}

	@Override
	public void onTable(TableCoordinates tableCoordinates) {

	}

}
