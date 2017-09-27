package org.wickedsource.docxstamper.processor;

import org.wickedsource.docxstamper.api.commentprocessor.ICommentProcessor;
import org.wickedsource.docxstamper.api.coordinates.ParagraphCoordinates;
import org.wickedsource.docxstamper.api.coordinates.RunCoordinates;
import org.wickedsource.docxstamper.api.coordinates.TableCoordinates;

public abstract class BaseCommentProcessor implements ICommentProcessor {

	private ParagraphCoordinates currentParagraphCoordinates;

	private RunCoordinates currentRunCoordinates;

	private CommentProcessorRegistry registry;

	public CommentProcessorRegistry getRegistry() {
		return registry;
	}

	@Override
	public void setRegistry(CommentProcessorRegistry registry) {
		this.registry = registry;
	}

	public RunCoordinates getCurrentRunCoordinates() {
		return currentRunCoordinates;
	}

	@Override
	public void setCurrentRunCoordinates(RunCoordinates currentRunCoordinates) {
		this.currentRunCoordinates = currentRunCoordinates;
	}

	@Override
	public void setCurrentParagraphCoordinates(ParagraphCoordinates coordinates) {
		this.currentParagraphCoordinates = coordinates;
	}

	public ParagraphCoordinates getCurrentParagraphCoordinates() {
		return currentParagraphCoordinates;
	}


	@Override
	public void onParagraphe(ParagraphCoordinates paragrapheCoordinates) {

	}

	@Override
	public void onTable(TableCoordinates tableCoordinates) {

	}

}
