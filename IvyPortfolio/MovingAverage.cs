//
// MovingAverage.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2016-2020 Jeffrey Stedfast
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using Newtonsoft.Json;

namespace IvyPortfolio
{
	public enum MovingAverage
	{
		Simple200Day  = 1,
		Simple10Month = 2,
		Simple12Month = 3
	}

	public class MovingAverageConverter : JsonConverter<MovingAverage>
	{
		public override MovingAverage ReadJson (JsonReader reader, Type objectType, MovingAverage existingValue, bool hasExistingValue, JsonSerializer serializer)
		{
			MovingAverage movingAverage;

			if (reader.Value == null)
				return 0;

			var value = reader.Value.ToString ();
			if (Enum.TryParse (value, out movingAverage))
				return movingAverage;

			switch (value.ToLowerInvariant ()) {
			case "200-day": return MovingAverage.Simple200Day;
			case "10-month": return MovingAverage.Simple10Month;
			case "12-month": return MovingAverage.Simple12Month;
			default: return 0;
			}
		}

		public override void WriteJson (JsonWriter writer, MovingAverage value, JsonSerializer serializer)
		{
			throw new NotImplementedException ();
		}
	}
}
