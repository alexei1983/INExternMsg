using System;
using System.Collections.Generic;
using System.Linq;

namespace org.goodspace.Utils.ImageNow {

  /// <summary>
  /// Provides functionality to read <see cref="INExternMsg"/> objects from the ImageNow database.
  /// </summary>
  public class INExternMsgReader {

    /// <summary>
    /// The ImageNow database connection string.
    /// </summary>
    public string ConnectionString { get; private set; }

    /// <summary>
    /// Whether or not to set the status to <see cref="ExternMsgStatus.Processing"/> when a 
    /// message is received.
    /// </summary>
    public bool SetStatusProcessingOnReceive { get; set; }

    /// <summary>
    /// Determines how <see cref="string"/> values are compared in an <see cref="INExternMsg"/> instance.
    /// </summary>
    public StringComparison StringComparison { get; set; } = StringComparison.InvariantCultureIgnoreCase;

    /// <summary>
    /// Event raised when an <see cref="INExternMsg"/> is received.
    /// </summary>
    public event EventHandler MessageReceived;

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgReader"/> class.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    public INExternMsgReader(string connectionString)
    {
      if (string.IsNullOrWhiteSpace(connectionString))
        throw new ArgumentException("Connection string is required.", nameof(connectionString));

      ConnectionString = connectionString;
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgReader"/> class.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <param name="setStatusProcessingOnReceive">Whether or not to set the status to <see cref="ExternMsgStatus.Processing"/> 
    /// when a message is received.</param>
    public INExternMsgReader(string connectionString,
                             bool setStatusProcessingOnReceive) : this(connectionString)
    {
      SetStatusProcessingOnReceive = setStatusProcessingOnReceive;
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsgReader"/> class.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <param name="setStatusProcessingOnReceive">Whether or not to set the status to <see cref="ExternMsgStatus.Processing"/> 
    /// when a message is received.</param>
    /// <param name="stringComparison">Determines how <see cref="string"/> values are compared in an 
    /// <see cref="INExternMsg"/> instance.</param>
    public INExternMsgReader(string connectionString,
                             bool setStatusProcessingOnReceive,
                             StringComparison stringComparison) : this(connectionString, setStatusProcessingOnReceive)
    {
      StringComparison = stringComparison;
    }

    /// <summary>
    /// Retrieves a single <see cref="INExternMsg"/> from the database using the specified 
    /// message ID.
    /// </summary>
    /// <param name="messageId">ID of the message to retrieve from the database.</param>
    /// <returns><see cref="INExternMsg"/> with the specified message ID.</returns>
    public INExternMsg Get(string messageId)
    {
      if (string.IsNullOrWhiteSpace(messageId))
        throw new ArgumentException("Message ID is required.", nameof(messageId));

      var message = INExternMsgHelper.GetMessages(ConnectionString,
                                                  INExternMsgHelper.IN_EXTERN_MSG_SELECT_ID_CMD,
                                                  new Dictionary<string, object>() { { "@MsgId", messageId } },
                                                  StringComparison);

      return message?.FirstOrDefault();
    }

    /// <summary>
    /// Retrieves all new, outbound <see cref="INExternMsg"/> objects of the specified 
    /// message type with the specified name from the database.
    /// </summary>
    /// <param name="messageType">The message type.</param>    
    /// <param name="messageName">The message name.</param>
    public void Receive(string messageType,
                        string messageName)
    {
      if (string.IsNullOrWhiteSpace(messageType))
        throw new ArgumentException("Message type is required.", nameof(messageType));

      if (string.IsNullOrWhiteSpace(messageName))
        throw new ArgumentException("Message name is required.", nameof(messageName));

      foreach (var message in INExternMsgHelper.GetMessages(ConnectionString,
                                                            INExternMsgHelper.IN_EXTERN_MSG_SELECT_CMD,
                                                            new Dictionary<string, object>() { { "@MsgName", messageName },
                                                                                               { "@MsgType", messageType }},
                                                            StringComparison))
      {
        if (SetStatusProcessingOnReceive)
          message.SetDatabaseStatus(ConnectionString, ExternMsgStatus.Processing);

        OnMessageReceived(message);
      }
    }

    /// <summary>
    /// Retrieves all new, outbound <see cref="INExternMsg"/> objects of the specified 
    /// message type from the database.
    /// </summary>
    /// <param name="messageType">The message type.</param>
    public void Receive(string messageType)
    {
      if (string.IsNullOrWhiteSpace(messageType))
        throw new ArgumentException("Message type is required.", nameof(messageType));

      foreach (var message in INExternMsgHelper.GetMessages(ConnectionString,
                                                            INExternMsgHelper.IN_EXTERN_MSG_SELECT_TYPE_CMD,
                                                            new Dictionary<string, object>() { { "@MsgType", messageType } },
                                                            StringComparison))
      {
        if (SetStatusProcessingOnReceive)
          message.SetDatabaseStatus(ConnectionString, ExternMsgStatus.Processing);

        OnMessageReceived(message);
      }
    }

    /// <summary>
    /// Retrieves all new, outbound <see cref="INExternMsg"/> objects from the database.
    /// </summary>
    public void Receive()
    {
      foreach (var message in INExternMsgHelper.GetMessages(ConnectionString,
                                                            INExternMsgHelper.IN_EXTERN_MSG_SELECT_ALL_CMD,
                                                            null,
                                                            StringComparison))
      {
        if (SetStatusProcessingOnReceive)
          message.SetDatabaseStatus(ConnectionString, ExternMsgStatus.Processing);

        OnMessageReceived(message);
      }
    }

    /// <summary>
    /// Raises the <see cref="MessageReceived"/> event.
    /// </summary>
    /// <param name="message">The <see cref="INExternMsg"/> object that was received.</param>
    protected void OnMessageReceived(INExternMsg message)
    {
      if (message != null)
        MessageReceived?.Invoke(message, EventArgs.Empty);
    }
  }
}
