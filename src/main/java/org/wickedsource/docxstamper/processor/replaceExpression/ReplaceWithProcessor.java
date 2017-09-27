package org.wickedsource.docxstamper.processor.replaceExpression;

import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.R;
import org.wickedsource.docxstamper.processor.BaseCommentProcessor;
import org.wickedsource.docxstamper.util.RunUtil;

public class ReplaceWithProcessor extends BaseCommentProcessor
		implements IReplaceWithProcessor {

	public Map<R, String> replaces = new HashMap<>();

	public ReplaceWithProcessor() {

	}

	@Override
	public void commitChanges(WordprocessingMLPackage document) {
		for (Entry<R, String> entry : replaces.entrySet()) {
			RunUtil.setText(entry.getKey(), entry.getValue());
		}
	}

	@Override
	public void reset() {
		replaces.clear();
	}

	@Override
	public void replaceWordWith(String expression) {
		if (this.getCurrentRunCoordinates() != null) {
			replaces.put(this.getCurrentRunCoordinates().getRun(), expression);
		}

	}

}
