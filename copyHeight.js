// Copy Height

let aw = host.ActiveSelectionRange.All();

for (let step = 2; step <= aw.Count; step++) {
	let aspectRatio = aw.Item(step).SizeWidth / aw.Item(step).SizeHeight;

	aw.Item(step).SizeHeight = aw.Item(1).SizeHeight;
	aw.Item(step).SizeWidth = aw.Item(step).SizeHeight * aspectRatio;
}
