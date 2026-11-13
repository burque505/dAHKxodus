// dAHKxodus's extension point (dAHKxodus.cs, through #load). In dAHKxodus it adds nothing.
//
// dAHKxodus Pro (its own, private repository) puts its own version of this file here: it #loads its files, and
// Started sets them up. Everything else stays as in dAHKxodus, so the two stay one code base and dAHKxodus's changes
// merge into Pro as they are (Pro keeps its own copy of this file: merge=ours).
static class DaxExtensions
{
	/// The window and its WebView2 are ready (the page may still be loading): an extension sets itself up here.
	public static void Started(MainForm form) { }
}
