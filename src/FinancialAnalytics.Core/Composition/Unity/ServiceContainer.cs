using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Practices.Unity;
using FinancialAnalytics.Core.Composition.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
    public class ServiceContainer : IServiceContainer
    {
        private readonly IUnityContainer _container;

        
        /// <summary>
        /// used to make unity container thread safe for registering
        /// </summary>
        private readonly object _containerLock = new object();

        public ServiceContainer() : this(new UnityContainer()) { }

        public ServiceContainer(IUnityContainer container)
        {
            _container = container;
            _container.RegisterType<IServiceContainer, ServiceContainer>(Lifetime.Singleton);
            _container.RegisterInstance<IServiceContainer>(this);
        }

        public T GetInstance<T>()
        {
            return _container.Resolve<T>();
        }

        public IEnumerable<T> GetAllInstances<T>()
        {
            return _container.ResolveAll<T>();
        }

        public T GetInstance<T>(string serviceName)
        {
            return _container.Resolve<T>(serviceName);
        }

        public TService GetInstance<TService>(params ResolverOverride[] overrides)
        {
            return _container.Resolve<TService>(overrides);
        }

        public object GetService(Type serviceType)
        {
            return _container.Resolve(serviceType);
        }

        public object GetInstance(Type serviceType)
        {
            return _container.Resolve(serviceType);
        }

        public IEnumerable<object> GetAllInstances(Type serviceType)
        {
            return _container.ResolveAll(serviceType);
        }

        public object GetInstance(Type serviceType, string serviceName)
        {
            return _container.Resolve(serviceType, serviceName);
        }

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(TInterface instance, LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                _container.RegisterInstance<TInterface>(instance, lifetimeManager);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(lifetimeManager);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>()
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>();
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(TInterface instance)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TInterface>(_container, instance.GetType());
                _container.RegisterInstance<TInterface>(instance);
            }
            return this;
        }
        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(string name)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(name);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(string name, LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(name, lifetimeManager);
            }
            return this;
        }


   

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(string name, TInterface instance)
        {
            lock (_containerLock)
            {
                _container.RegisterInstance(name, instance);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<T>(LifetimeManager lifetimeManager)
        {
            
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<T>(_container, typeof(T));
                _container.RegisterType<T>(lifetimeManager);
            }
            return this;
        }

        public IServiceContainer RegisterType<T>(LifetimeManager lifetimeManager, Func<IServiceLocator, T> factory)
        {
 
            lock (_containerLock)
            {
               // ThrowExceptionIfRegistered<T>(typeof(T));
                _container.RegisterType<T>(lifetimeManager, new InjectionFactory(x => factory(new ServiceLocator(_container))));
            }
            return this;
        }

        public IServiceContainer RegisterType<T>(Func<IServiceLocator, T> factory)
        {

            lock (_containerLock)
            {
               // ThrowExceptionIfRegistered<T>(typeof(T));
                _container.RegisterType<T>(new InjectionFactory(x => factory(new ServiceLocator(_container))));
            }
            return this;
        }

        public void Teardown(object o)
        {
            _container.Teardown(o);
        }

        IServiceContainer IServiceContainer.Configure(params IConfigureContainer[] containerConfigurators)
        {
            foreach (var item in containerConfigurators)
            {
                item.ConfigureContainer(this);
            }
            return this;
        }

    }
}
