import {
  Logger,
  NotificationTypes,
} from "@medusajs/framework/types";
import {
  AbstractNotificationProviderService,
  MedusaError,
} from "@medusajs/framework/utils";
import {
  EmailClient,
  KnownEmailSendStatus,
} from "@azure/communication-email";

type InjectedDependencies = {
  logger: Logger
}

type AzureCommunicationServicesOptions = {
  connectionString: string
  defaultFromAddress: string
  defaultNotificationSubject: string
}

type AzureCommunicationServicesConfig = {
  defaultFromAddress: string
  defaultNotificationSubject: string
}

class AzureCommunicationServicesNotificationProviderService extends AbstractNotificationProviderService {
  static identifier = "azure-notification"
  protected config_: AzureCommunicationServicesConfig
  protected logger_: Logger
  protected client_: EmailClient

  constructor({ logger }: InjectedDependencies, options: AzureCommunicationServicesOptions) {
    super()

    this.config_ = {
      defaultFromAddress: options.defaultFromAddress,
      defaultNotificationSubject: options.defaultNotificationSubject,
    }
    this.logger_ = logger

    this.client_ = new EmailClient(options.connectionString)
  }

  static validateOptions(options: AzureCommunicationServicesOptions) {
    if (!options.connectionString) {
      throw new MedusaError(
        MedusaError.Types.INVALID_DATA,
        "`connectionString` is required in the provider's options."
      )
    }

    if (!options.defaultFromAddress) {
      throw new MedusaError(
        MedusaError.Types.INVALID_DATA,
        "`defaultFromAddress` is required in the provider's options."
      )
    }

    if (!options.defaultNotificationSubject) {
      throw new MedusaError(
        MedusaError.Types.INVALID_DATA,
        "`defaultNotificationSubject` is required in the provider's options."
      )
    }
  }

  async send({ channel,
    from,
    to,
    content: notificationContent,
    attachments: notificationAttachments
  }: NotificationTypes.ProviderSendNotificationDTO): Promise<NotificationTypes.ProviderSendNotificationResultsDTO> {
    if (channel !== "email") {
      throw new MedusaError(
        MedusaError.Types.INVALID_ARGUMENT,
        `Unsupported channel: ${channel}. Only 'email' is supported.`
      )
    }

    const senderAddress = from?.trim() || this.config_.defaultFromAddress

    const recipients = {
      to: [{ address: to }],
    }

    const content = {
      subject: notificationContent?.subject ?? this.config_.defaultNotificationSubject,
      plainText: notificationContent?.text ?? (notificationContent?.html ? notificationContent.html.replace(/<[^>]*>/g, "") : ""),
      html: notificationContent?.html,
    }

    let attachments;

    if (Array.isArray(notificationAttachments)) {
      attachments = notificationAttachments.map((attachment) => ({
        name: attachment.filename,
        contentId: attachment.disposition == "inline" ? attachment.id : undefined,
        contentType: attachment.content_type ?? "application/octet-stream", // MIME type (e.g., 'application/pdf')
        contentInBase64: attachment.content, // base64 encoded string of the file
      }))
    }

    const message = {
      senderAddress,
      recipients,
      content,
      attachments,
    }

    try {
      this.logger_.info(`Sending email to: ${to}`);

      const poller = await this.client_.beginSend(message);

      const result = await poller.pollUntilDone();

      if (result.status !== KnownEmailSendStatus.Succeeded) {
        throw new Error(result.error?.message ?? `Email send failed with status: ${result.status}`);
      }

      this.logger_.info(`Email sent successfully (operation id: ${result.id})`);

      return { id: result.id };
    }
    catch (e) {
      this.logger_.error(e instanceof Error ? e : String(e))
      throw new Error(`Failed to send email: ${e instanceof Error ? e.message : e}`)
    }
  }
}

export default AzureCommunicationServicesNotificationProviderService