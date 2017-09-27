package org.wickedsource.docxstamper.processor;

import org.wickedsource.docxstamper.api.coordinates.ParagraphCoordinates;
import org.wickedsource.docxstamper.api.coordinates.RunCoordinates;
import org.wickedsource.docxstamper.util.CommentWrapper;

public interface ICommentProcessorRegistry {
	public interface Info {
		CommentWrapper getCommentWrapper();
		String getComment();
		ParagraphCoordinates getParagrapheCoordinates();
		RunCoordinates getRunCoordinates();
	}
}
