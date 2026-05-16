// Licensed under the Apache License, Version 2.0
// Copyright 2025, Mindful Software LLC, All rights reserved.

import 'package:dartastic_opentelemetry/dartastic_opentelemetry.dart';

import 'otel_outlook_calendar_suppression.dart';

const _tracerName = 'dartastic_otel_libs_outlook_calendar';

Tracer _tracer() => OTel.tracerProvider().getTracer(_tracerName);

class OutlookCalendarResult {
  const OutlookCalendarResult({
    this.itemCount,
    this.attendeeCount,
    this.usedDeltaToken,
    this.deltaTokenReturned,
    this.deltaLinkExpired,
    this.onlineMeetingCreated,
    this.onlineMeetingProvider,
    this.confidence,
    this.eventId,
    this.requestId,
    this.statusCode,
  });

  /// Items returned (`value.length` for collection endpoints).
  final int? itemCount;

  /// Number of attendees in findMeetingTimes / event create — affects
  /// Graph throttling tier.
  final int? attendeeCount;

  /// Whether the call used a `@odata.deltaLink` (incremental).
  final bool? usedDeltaToken;

  /// New `@odata.deltaLink` returned (telemetry, not the URL).
  final bool? deltaTokenReturned;

  /// `410 Gone` — delta link expired, full resync required (incident
  /// signal).
  final bool? deltaLinkExpired;

  /// True if Teams meeting was provisioned.
  final bool? onlineMeetingCreated;

  /// `teamsForBusiness` / `skypeForBusiness` / `skypeForConsumer`.
  final String? onlineMeetingProvider;

  /// `meetingTimeSuggestions[].confidence` (high/medium/low) for
  /// findMeetingTimes — surface chosen result.
  final String? confidence;

  /// Event id (`id` field).
  final String? eventId;

  /// `request-id` header.
  final String? requestId;

  final int? statusCode;
}

class OutlookCalendarCallResult<T> {
  const OutlookCalendarCallResult({required this.value, this.result});
  final T value;
  final OutlookCalendarResult? result;
}

/// Wraps a Microsoft Graph calendar call in a CLIENT span. [operation]
/// is one of `events.list`, `events.get`, `events.create`,
/// `events.update`, `events.delete`, `events.instances`,
/// `events.accept`, `events.decline`, `events.tentativelyAccept`,
/// `calendarView`, `delta`, `findMeetingTimes`, `getSchedule`,
/// `subscriptions.create`.
Future<T> tracedOutlookCalendar<T>({
  required String operation,
  required Future<OutlookCalendarCallResult<T>> Function() invoke,
  String? mailbox,
  String? calendarId,
}) async {
  if (outlookCalendarInstrumentationSuppressed()) {
    final r = await invoke();
    return r.value;
  }
  final m = <String, Object>{
    'rpc.system': 'outlook_calendar',
    'rpc.method': operation,
    'outlook_calendar.operation': operation,
  };
  if (mailbox != null) m['outlook_calendar.mailbox'] = mailbox;
  if (calendarId != null) m['outlook_calendar.calendar.id'] = calendarId;

  final span = _tracer().startSpan(
    'outlook_calendar $operation',
    kind: SpanKind.client,
    attributes: OTel.attributesFromMap(m),
  );
  try {
    final r = await invoke();
    if (r.result != null) {
      final extras = <Attribute<Object>>[];
      if (r.result!.itemCount != null) {
        extras.add(OTel.attributeInt(
          'outlook_calendar.item_count',
          r.result!.itemCount!,
        ));
      }
      if (r.result!.attendeeCount != null) {
        extras.add(OTel.attributeInt(
          'outlook_calendar.attendee_count',
          r.result!.attendeeCount!,
        ));
      }
      if (r.result!.usedDeltaToken != null) {
        extras.add(OTel.attributeBool(
          'outlook_calendar.delta_token.used',
          r.result!.usedDeltaToken!,
        ));
      }
      if (r.result!.deltaTokenReturned != null) {
        extras.add(OTel.attributeBool(
          'outlook_calendar.delta_token.returned',
          r.result!.deltaTokenReturned!,
        ));
      }
      if (r.result!.deltaLinkExpired != null) {
        extras.add(OTel.attributeBool(
          'outlook_calendar.delta_link.expired',
          r.result!.deltaLinkExpired!,
        ));
      }
      if (r.result!.onlineMeetingCreated != null) {
        extras.add(OTel.attributeBool(
          'outlook_calendar.online_meeting.created',
          r.result!.onlineMeetingCreated!,
        ));
      }
      if (r.result!.onlineMeetingProvider != null) {
        extras.add(OTel.attributeString(
          'outlook_calendar.online_meeting.provider',
          r.result!.onlineMeetingProvider!,
        ));
      }
      if (r.result!.confidence != null) {
        extras.add(OTel.attributeString(
          'outlook_calendar.meeting_time.confidence',
          r.result!.confidence!,
        ));
      }
      if (r.result!.eventId != null) {
        extras.add(OTel.attributeString(
          'outlook_calendar.event.id',
          r.result!.eventId!,
        ));
      }
      if (r.result!.requestId != null) {
        extras.add(OTel.attributeString(
          'outlook_calendar.request.id',
          r.result!.requestId!,
        ));
      }
      if (r.result!.statusCode != null) {
        extras.add(OTel.attributeInt(
          'http.response.status_code',
          r.result!.statusCode!,
        ));
        if (r.result!.statusCode! >= 400) {
          span.setStatus(
            SpanStatusCode.Error,
            'Outlook Calendar HTTP ${r.result!.statusCode}',
          );
        }
      }
      if (extras.isNotEmpty) span.addAttributes(OTel.attributes(extras));
    }
    return r.value;
  } catch (e, st) {
    span.addAttributes(OTel.attributes([
      OTel.attributeString(
        ErrorResource.errorType.key,
        e.runtimeType.toString(),
      ),
    ]));
    span.recordException(e, stackTrace: st);
    span.setStatus(SpanStatusCode.Error, e.toString());
    rethrow;
  } finally {
    span.end();
  }
}
