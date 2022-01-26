// Rename with dimension

host.ActivePage.FindShapes(undefined, cdrRectangleShape, true).SetPixelAlignedRendering(true)

let aw = host.ActivePage.Shapes.All();
for (let step = 1; step <= aw.Count; step++) {
	let pixel = 300;
	let item = aw.Item(step);
	let width = Math.round(item.SizeWidth * pixel);
	let height = Math.round(item.SizeHeight * pixel);
	item.Name = `${width}x${height}`;
}
