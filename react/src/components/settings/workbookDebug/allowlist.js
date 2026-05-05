// Gate the Workbook Debug "Reset workbook defaults" action to the
// workbook's author. We compare the Office workbook `author` property
// against the signed-in user's SSO name; only an exact match (after
// trim + lowercase) returns true.
//
// This is a guardrail against accidents, not a security boundary — the
// Excel author property is editable by anyone with write access.

export function isAuthorMatch(author, currentUserName) {
	const a = String(author ?? '').trim().toLowerCase();
	const u = String(currentUserName ?? '').trim().toLowerCase();
	if (!a || !u) return false;
	return a === u;
}
