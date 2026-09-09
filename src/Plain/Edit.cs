namespace Plain;

/// <summary>
/// One step of undo, and where it can be worked out, how to do the thing again.
///
/// Not every edit can say how to repeat itself, and this is deliberately honest about that: a step with no Redo
/// can be undone but not redone, and undoing one throws away anything waiting to be redone rather than leaving a
/// stack that would put things back in the wrong order. Half a redo is worse than none.
///
/// There is an implicit conversion from a plain Action so that anything which only knows how to undo can carry on
/// saying so, without every place that makes an edit having to think about redo.
/// </summary>
public sealed record Edit(Action Undo, Action? Redo)
{
    public static implicit operator Edit(Action undo) => new(undo, null);

    public bool CanRedo => Redo is not null;
}
