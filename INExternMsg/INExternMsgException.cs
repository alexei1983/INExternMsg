using System;

namespace org.goodspace.Utils.ImageNow {

  /// <summary>
  /// Represents an <see cref="Exception"/> that occurred in an <see cref="INExternMsg"/> object.
  /// </summary>
  public class INExternMsgException : Exception {

    /// <summary>
    /// The <see cref="INExternMsg"/> ID associated with the object that caused the exception.
    /// </summary>
    public string MessageId { get; set; }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgException"/> class.
    /// </summary>
    /// <param name="message">The exception message.</param>
    /// <param name="innerException">The inner <see cref="Exception"/>.</param>
    public INExternMsgException(string message, Exception innerException) : base(message, innerException)
    {
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgException"/> class.
    /// </summary>
    /// <param name="message">The exception message.</param>
    public INExternMsgException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgException"/> class.
    /// </summary>
    /// <param name="message">The exception message.</param>
    /// <param name="inExternMsgId">The <see cref="INExternMsg"/> ID associated with the 
    /// object that caused the exception.</param>
    public INExternMsgException(string message, string inExternMsgId) : base(message)
    {
      MessageId = inExternMsgId;
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgException"/> class.
    /// </summary>
    /// <param name="message">The exception message.</param>
    /// <param name="inExternMsgId">The <see cref="INExternMsg"/> ID associated with the 
    /// object that caused the exception.</param>
    /// <param name="innerException">The inner <see cref="Exception"/>.</param>
    public INExternMsgException(string message,
                                string inExternMsgId,
                                Exception innerException) : base(message, innerException)
    {
      MessageId = inExternMsgId;
    }
  }
}
