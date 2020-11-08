using System;

namespace org.goodspace.Utils.ImageNow {

  /// <summary>
  /// Represents event data associated with adding a property to an <see cref="INExternMsg"/>.
  /// </summary>
  public class PropertyAddedEventArgs : EventArgs {

    /// <summary>
    /// The name of the property that was added.
    /// </summary>
    public string PropertyName { get; set; }

    /// <summary>
    /// The value of the property that was added.
    /// </summary>
    public string PropertyValue { get; set; }
  }
}
