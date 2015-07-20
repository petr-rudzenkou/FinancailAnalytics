using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Practices.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
    internal class UnityContainerChecker
    {
        /// <summary>
        /// Used during development to prevent duplicated registrations which has performance impact and can lead to subtle errors with stateful classes 
        /// (e.g. first registration resolved with some side effect, then second registration done, then second used with repeats action not expected to be repeated)
        /// </summary>
        /// <typeparam name="TFrom"></typeparam>
        /// <param name="container"></param>
        /// <param name="typeToCheck"></param>
        [DebuggerStepThrough]
        [DebuggerStepperBoundary]
        [Conditional("DEBUG")]
        public static void ThrowExceptionIfRegistered<TFrom>(IUnityContainer container, Type typeToCheck, string nameToCheck = null)
        {
            //see details in https://unity.codeplex.com/discussions/268223
            ContainerRegistration[] registrations = null;
            while (registrations == null)
            {
                try
                {
                    registrations = container.Registrations.ToArray();
                }
                catch (InvalidOperationException)
                {
                    // good enough for DEBUG to fix "Collection was modified; enumeration operation may not execute."
                    // happens when collection changed during resolving, all registrations are on locks already
                    // avoids adding locks for resolving into RELEASE code
                }
            }

            ContainerRegistration alreadyHere = registrations.SingleOrDefault(x => x.RegisteredType == typeof(TFrom) && x.Name == nameToCheck);
            if (alreadyHere != null)
            {
                var same = typeToCheck == alreadyHere.MappedToType ? " (the same)" : " (different)";
                throw new InvalidOperationException(typeof(TFrom) + " already registered and implemented by " + alreadyHere + same + ".");
            }
        }


    }
}
