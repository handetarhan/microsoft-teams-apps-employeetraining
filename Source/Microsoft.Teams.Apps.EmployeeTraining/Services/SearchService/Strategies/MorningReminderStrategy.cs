// <copyright file="MorningReminderStrategy.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

// 19.10.2021 smarttek
namespace Microsoft.Teams.Apps.EmployeeTraining.Services.SearchService.Strategies
{
    using System;
    using System.Globalization;
    using Microsoft.Teams.Apps.EmployeeTraining.Models;

    /// <summary>
    /// Generates filter query for fetching events to send morning notifications.
    /// </summary>
    public class MorningReminderStrategy : IFilterGeneratingStrategy
    {
        /// <inheritdoc/>
        public string GenerateFilterQuery(SearchParametersDto searchParametersDto)
        {
            var eventDate = DateTime.UtcNow.Date;

            return $"{nameof(EventEntity.Status)} eq {(int)EventStatus.Active} and " +
                $"{nameof(EventEntity.StartDate)} eq {eventDate.ToString("O", CultureInfo.InvariantCulture)} and " +
                $"{nameof(EventEntity.RegisteredAttendeesCount)} gt 0";
        }
    }
}