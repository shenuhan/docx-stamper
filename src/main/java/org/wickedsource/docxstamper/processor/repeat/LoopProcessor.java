package org.wickedsource.docxstamper.processor.repeat;

import java.util.ArrayList;
import java.util.List;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.wickedsource.docxstamper.api.coordinates.AbstractCoordinates;
import org.wickedsource.docxstamper.api.coordinates.ParagraphCoordinates;
import org.wickedsource.docxstamper.api.coordinates.TableCoordinates;
import org.wickedsource.docxstamper.processor.BaseCommentProcessor;

public class LoopProcessor extends BaseCommentProcessor implements ILoopProcessor {
	private static class Loop {
		private int indexStart;
		private int indexEnd;
		private String name;
		private List<Object> objects;
		private List<AbstractCoordinates> coordinates = new ArrayList<>();
	}

	private List<Loop> loops = new ArrayList<>();
	private List<Loop> endedLoops = new ArrayList<>();


	@Override
	public void loop(List<Object> objects, String name) {
		Loop loop = new Loop();
		loop.indexStart = getCurrentParagraphCoordinates().getIndex();
		loop.name = name;
		loop.objects = objects;
		loops.add(loop);
		getRegistry().getExpressionResolver().addVariable(name, objects.get(0));
	}

	@Override
	public void endloop() {
		Loop loop = loops.get(loops.size()-1);
		loop.indexEnd = getCurrentParagraphCoordinates().getIndex();
		getRegistry().getExpressionResolver().removeVariable(loop.name);
		endedLoops.add(loop);
		loops.remove(loops.size()-1);
	}

	@Override
	public void commitChanges(WordprocessingMLPackage document) {

	}

	@Override
	public void reset() {
		loops.clear();
	}


	@Override
	public void onParagraphe(ParagraphCoordinates paragrapheCoordinates) {
		if (loops.isEmpty()) return;

		Loop loop = loops.get(loops.size()-1);
		loop.coordinates.add(paragrapheCoordinates);
	}

	@Override
	public void onTable(TableCoordinates tableCoordinates) {
		if (loops.isEmpty()) return;

		Loop loop = loops.get(loops.size()-1);
		loop.coordinates.add(tableCoordinates);
	}

}
