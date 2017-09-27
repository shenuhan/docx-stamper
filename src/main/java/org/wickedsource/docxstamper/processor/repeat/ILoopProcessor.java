package org.wickedsource.docxstamper.processor.repeat;

import java.util.List;

public interface ILoopProcessor {
	void loop(List<Object> objects, String name);
	void endloop();
}
